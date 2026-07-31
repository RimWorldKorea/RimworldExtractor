using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore.DefTreeSimulator
{
    /// <summary>
    /// LoadFolders 및 조건부 PatchOperation이 유발하는 다형적 로드 상황의 시뮬레이션 분기 생성을 전담하는 클래스입니다.
    /// </summary>
    public static class DoctorStrange
    {
        /// <summary>
        /// LoadFolders.xml의 조건 조합(Any/All)을 계산하여 초기 스냅샷(멀티버스)들을 쫙 펼칩니다.
        /// </summary>
        public static List<DefSnapshot> GenerateInitialSnapshots(XDocument prePiledTree, List<ExtractableFolder> allSelectedFolders)
        {
            var snapshots = new List<DefSnapshot>();

            // 1. 모든 가능한 모드 로드 경우의 수(PID Combinations) 추출
            List<string[]> allPossibleUniverses = CalculateAllModCombinations(allSelectedFolders);

            // 2. 각 경우의 수(우주)마다 스냅샷을 생성하고, 조건에 맞는 폴더를 할당
            foreach (var modCombination in allPossibleUniverses)
            {
                var snapshot = new DefSnapshot(new XDocument(prePiledTree), modCombination);

                // 3. 이 우주(조합)에서 활성화될 수 있는 폴더만 필터링하여 할당
                foreach (var folder in allSelectedFolders)
                {
                    if (IsFolderActiveInThisUniverse(folder, modCombination))
                    {
                        if (folder.FolderName.EndsWith("Patches", StringComparison.OrdinalIgnoreCase))
                            snapshot.AssignedPatches.Add(folder);
                        else
                            snapshot.AssignedFolders.Add(folder);
                    }
                }

                snapshots.Add(snapshot);
            }

            return snapshots;
        }

        /// <summary>
        /// RMK의 LoadFoldersBuilder 로직으로부터, 폴더들의 로드 조건으로부터 파생되는 
        /// 모든 유니버스(활성화된 모드 조합)의 경우의 수를 계산합니다.
        /// </summary>
        private static List<string[]> CalculateAllModCombinations(List<ExtractableFolder> folders)
        {
            var combinations = new HashSet<string[]>(new StringArrayComparer());
            
            // Base universe (어떤 조건부 모드도 활성화되지 않은 기본 상태)
            combinations.Add(Array.Empty<string>()); 

            foreach (var folder in folders)
            {
                if (string.IsNullOrEmpty(folder.RequiredPackageId)) continue;

                var pids = folder.RequiredPackageId.Split(',')
                    .Select(p => p.Trim())
                    .Where(p => !string.IsNullOrEmpty(p))
                    .ToArray();

                // RMK의 IfModActive(Any) 처리 논리에 따라, 
                // 콤마로 구분된 각 PackageID를 각각 독립적인 유니버스로 분화시킵니다.
                foreach (var pid in pids)
                {
                    combinations.Add(new[] { pid });
                }

                // 추후 IfModActiveAll(All)처럼 전체가 다 활성화되어야 하는 
                // 복합 조건을 대비해 전체 묶음(PID 전체 교집합)도 하나의 유니버스로 등록합니다.
                if (pids.Length > 1)
                {
                    combinations.Add(pids);
                }
            }
            return combinations.ToList();
        }

        /// <summary>
        /// 특정 유니버스(활성화된 모드 목록)에서 해당 폴더가 로드(True)되어야 하는지 평가합니다.
        /// </summary>
        private static bool IsFolderActiveInThisUniverse(ExtractableFolder folder, string[] activeModsInUniverse)
        {
            if (string.IsNullOrEmpty(folder.RequiredPackageId)) return true;

            var pids = folder.RequiredPackageId.Split(',')
                .Select(p => p.Trim())
                .Where(p => !string.IsNullOrEmpty(p))
                .ToArray();

            // 현재 ModLister는 IfModActive 속성을 추출하므로 RMK의 BindingMode.Any 에 해당합니다.
            // 따라서 요구되는 패키지 ID 중 '하나라도' 현재 유니버스에 존재하면(활성화되면) 폴더를 로드합니다.
            return pids.Any(pid => activeModsInUniverse.Contains(pid, StringComparer.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 패치 폴더(Patches)의 XML을 읽어와 PatchOperation을 수행하고, 
        /// FindMod에 의해 분기된 유니버스(DefSnapshot)들을 반환합니다.
        /// </summary>
        public static List<DefSnapshot> ProcessPatchOperations(DefSnapshot currentUniverse, List<ExtractableFolder> patches)
        {
            var resultingUniverses = new List<DefSnapshot> { currentUniverse };

            foreach (var patchFolder in patches)
            {
                var folderPath = patchFolder.FullPath;
                if (!Directory.Exists(folderPath)) continue;

                // Patches 폴더 안의 모든 .xml 파일을 탐색합니다.
                var xmlFiles = FileInterface.DescendantFiles(folderPath)
                    .Where(x => x.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var doc = FileInterface.ReadXml(xmlFile);
                        var root = doc.Root;

                        // 림월드 패치 파일은 일반적으로 <Patch> 루트 아래에 <Operation> 노드들이 존재합니다.
                        if (root != null && root.Name.LocalName == "Patch")
                        {
                            var operations = root.Elements("Operation");
                            foreach (var operationNode in operations)
                            {
                                var nextUniverses = new List<DefSnapshot>();
                                
                                // 현재 존재하는 모든 유니버스에 해당 오퍼레이션을 적용해봅니다.
                                foreach (var universe in resultingUniverses)
                                {
                                    // ApplyPatchRecursive 내에서 FindMod를 만나면 
                                    // True/False 세계선이 갈라지며 nextUniverses가 늘어날 수 있습니다.
                                    nextUniverses.AddRange(PatchOperations.ApplyPatchRecursive(operationNode, universe));
                                }
                                
                                resultingUniverses = nextUniverses;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Wrn($"패치 XML 파일 처리 중 오류 발생 ({xmlFile}): {e.Message}");
                    }
                }
            }

            return resultingUniverses;
        }

        // 중복된 문자열 배열(조합)을 걸러내기 위한 비교기
        private class StringArrayComparer : IEqualityComparer<string[]>
        {
            public bool Equals(string[]? x, string[]? y)
            {
                if (x == y) return true;
                if (x == null || y == null) return false;
                if (x.Length != y.Length) return false;
                
                return !x.Except(y, StringComparer.OrdinalIgnoreCase).Any();
            }
            public int GetHashCode(string[] obj)
            {
                int hash = 17;
                foreach (var str in obj.OrderBy(s => s, StringComparer.OrdinalIgnoreCase))
                    hash = hash * 31 + StringComparer.OrdinalIgnoreCase.GetHashCode(str);
                return hash;
            }
        }
    }
}