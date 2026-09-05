using System.Xml.Linq;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore.DefTreeSimulator;

/// <summary>
/// LoadFolders 및 조건부 PatchOperation이 유발하는 다형적 로드 상황의 시뮬레이션 분기 생성을 전담하는 클래스입니다.
/// </summary>
public static class DoctorStrange
{
    /// <summary>
    /// 대상 모드들의 LoadFolders.xml (Any/All) 규칙을 분석하여 초기 평행우주 스냅샷들을 생성합니다.
    /// </summary>
    public static List<DefSnapshot> GenerateInitialSnapshots(XDocument prePiledTree, IEnumerable<Mod> targetMods, Version targetVersion)
    {
        var snapshots = new List<DefSnapshot>();
        
        // 1. 발생 가능한 모든 PackageID 조합(멀티버스)을 계산
        List<string[]> allPossibleUniverses = CalculateAllModCombinations(targetMods, targetVersion);
        
        Log.Common($"[멀티버스] 총 {allPossibleUniverses.Count}개의 초기 우주가 계산되었습니다.");

        foreach (var modCombination in allPossibleUniverses)
        {
            var activeModsSet = new HashSet<string>(modCombination, StringComparer.OrdinalIgnoreCase);
            
            var snapshot = new DefSnapshot(new XDocument(prePiledTree), modCombination);
            if (modCombination.Length == 0) snapshot.IsBaseSnapshot = true;

            // 2. 해당 우주에서 각 모드가 활성화해야 할 폴더 상대 경로 산출
            foreach (var mod in targetMods)
            {
                var activePaths = EvaluateActivePaths(mod, targetVersion, activeModsSet);
                
                foreach (var path in activePaths)
                {
                    // [핵심] 이미 생성된 FolderLoadUnit이 있는지 OwnedFolders에서 검색하여 재사용
                    var unit = mod.OwnedFolders.FirstOrDefault(u => string.Equals(u.RelativePath, path, StringComparison.OrdinalIgnoreCase));
                    if (unit == null)
                    {
                        // 없다면 새로 생성 (생성자 내부에서 자동으로 mod.OwnedFolders에 등록됨)
                        unit = new FolderLoadUnit(mod, path); 
                    }
                    
                    snapshot.LoadedUnits.Add(unit);
                }
            }
            
            Log.Common($"[유니버스 생성] 활성 조건: {(modCombination.Length > 0 ? string.Join(", ", snapshot.RequiredModIds) : "기본 우주(Base)")}");
            snapshots.Add(snapshot);
        }

        return snapshots;
    }

    /// <summary>
    /// 모드들의 LoadFoldersRules에 명시된 IfModActive, IfModActiveAll 조건을 바탕으로 모든 경우의 수를 계산합니다.
    /// </summary>
    private static List<string[]> CalculateAllModCombinations(IEnumerable<Mod> mods, Version targetVersion)
    {
        var combinations = new HashSet<string[]>(new StringArrayComparer());
        combinations.Add(Array.Empty<string>()); // 기본 우주(Base universe)

        foreach (var mod in mods)
        {
            if (mod.LoadFolders != null && mod.LoadFolders.Rules.TryGetValue(targetVersion, out var rules))
            {
                // RMK: IfModActive (Any) - 각 조건 모드가 독립적인 멀티버스 분기 생성
                foreach (var conditionSet in rules.ifActive.Keys)
                {
                    foreach (var pid in conditionSet)
                    {
                        combinations.Add(new[] { pid });
                    }
                }

                // RMK: IfModActiveAll (All) - 조건 모드 전체 묶음이 하나의 멀티버스 분기 생성
                foreach (var conditionSet in rules.ifActiveAll.Keys)
                {
                    if (conditionSet.Count > 1)
                    {
                        combinations.Add(conditionSet.ToArray());
                    }
                }
            }
        }
        return combinations.OrderBy(x => x.Length).ToList();
    }

    /// <summary>
    /// 특정 모드의 특정 멀티버스 조합에서 실제로 로드되어야 할 폴더들의 상대 경로(문자열)를 평가하여 반환합니다.
    /// </summary>
    internal static IEnumerable<string> EvaluateActivePaths(Mod mod, Version targetVersion, HashSet<string> activeModsSet)
    {
        var resolvedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 1. LoadFolders.xml 규칙 적용
        if (mod.LoadFolders != null && mod.LoadFolders.Rules.TryGetValue(targetVersion, out var rules))
        {
            foreach (var folder in rules.defaultLoad) resolvedPaths.Add(folder);

            foreach (var (conditionSet, folders) in rules.ifActive)
                if (conditionSet.Overlaps(activeModsSet)) 
                    foreach (var folder in folders) resolvedPaths.Add(folder);

            foreach (var (conditionSet, folders) in rules.ifActiveAll)
                if (conditionSet.IsSubsetOf(activeModsSet)) 
                    foreach (var folder in folders) resolvedPaths.Add(folder);
        }
        // 2. LoadFolders.xml이 없는 경우 림월드 Fallback 로직 적용
        else if (!string.IsNullOrWhiteSpace(mod.Path))
        {
            string versionStr = targetVersion.ToString();
            if (Directory.Exists(Path.Combine(mod.Path, versionStr)))
            {
                resolvedPaths.Add(versionStr);
                if (Directory.Exists(Path.Combine(mod.Path, "Common")))
                    resolvedPaths.Add("Common");
            }
            else 
            {
                resolvedPaths.Add(""); // 루트 폴더("/")
            }
        }
        return resolvedPaths;
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
    /// 해당 우주에 로드된 폴더 유닛들 중 'Patches' 폴더가 존재하는 유닛들을 대상으로 PatchOperation을 수행합니다.
    /// </summary>
    public static List<DefSnapshot> ProcessPatchOperations(DefSnapshot currentUniverse)
    {
        var resultingUniverses = new List<DefSnapshot> { currentUniverse };

        foreach (var unit in currentUniverse.LoadedUnits)
        {
            // 경로 문자열 기반이 아닌, FolderLoadUnit의 디렉토리 프로퍼티를 직접 사용
            var patchDir = unit.Patches;
            if (patchDir == null || !patchDir.Exists) continue;

            var xmlFiles = patchDir.GetFiles("*.xml", SearchOption.AllDirectories);
            foreach (var xmlFile in xmlFiles)
            {
                try
                {
                    var doc = FileInterface.ReadXml(xmlFile.FullName);
                    var root = doc.Root;

                    if (root != null && root.Name.LocalName == "Patch")
                    {
                        var operations = root.Elements("Operation");
                        foreach (var operationNode in operations)
                        {
                            var nextUniverses = new List<DefSnapshot>();
                            
                            foreach (var universe in resultingUniverses)
                            {
                                nextUniverses.AddRange(PatchOperations.ApplyPatchRecursive(operationNode, universe));
                            }
                            resultingUniverses = nextUniverses;
                        }
                    }
                }
                catch (Exception e)
                {
                    Log.Wrn($"패치 XML 분석 중 오류 발생 ({xmlFile.FullName}): {e.Message}");
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