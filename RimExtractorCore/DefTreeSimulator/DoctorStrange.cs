using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore.DefTreeSimulator
{
    /// <summary>
    /// 다중 우주(Multiverse) 분기 생성을 전담하는 마법사 클래스입니다.
    /// LoadFolders 조합 및 조건부 PatchOperation에 의한 평행 우주 스냅샷 분열을 담당합니다.
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
        /// RMK의 PIDCombinations 계산 로직을 기반으로
        /// 가능한 모든 '모드 조합(우주)'의 리스트를 반환합니다.
        /// </summary>
        private static List<string[]> CalculateAllModCombinations(List<ExtractableFolder> folders)
        {
            var combinations = new HashSet<string[]>(new StringArrayComparer());
            
            // Base 우주 (아무 조건 없는 기본 상태)
            combinations.Add(Array.Empty<string>()); 

            foreach (var folder in folders)
            {
                // TODO: RMK의 LoadFoldersBuilder.CalculateCombinations 로직 이식 구역
            }

            return combinations.ToList();
        }

        /// <summary>
        /// 특정 폴더의 LoadFolders 조건이 현재 스냅샷의 모드 조합(우주)에서 참(True)인지 검사합니다.
        /// </summary>
        private static bool IsFolderActiveInThisUniverse(ExtractableFolder folder, string[] activeModsInUniverse)
        {
            if (folder.RequiredPackageId == null) return true;

            // TODO: 실제 ExtractableFolder의 조건 검사 로직 적용 (Any / All 판별)
            return true; 
        }

        /// <summary>
        /// 패치(Patches) 파일들을 순회하며 조건부 패치 분기(DefSnapshot)를 생성합니다. (2차 분열)
        /// </summary>
        public static List<DefSnapshot> ProcessPatchOperations(DefSnapshot currentUniverse, List<ExtractableFolder> patches)
        {
            var resultingUniverses = new List<DefSnapshot> { currentUniverse };

            // TODO: patches 폴더 내의 XML을 순회하며 PatchOperationFindMod 등을 만났을 때
            // currentUniverse.Clone()을 호출하여 resultingUniverses에 다중 우주를 추가하는 로직 구현

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