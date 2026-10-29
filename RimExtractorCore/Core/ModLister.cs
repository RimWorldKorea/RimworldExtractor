using RimExtractorCore.DataTypes;

namespace RimExtractorCore;

/// <summary>모드 정보 통합 관리자입니다.</summary>
public static class ModLister
{
    /// <summary>프로그램이 관리하는 모든 모드가 등록됩니다.</summary>
    private static readonly Dictionary<string, Mod> ModRegistry = new(StringComparer.OrdinalIgnoreCase);
    
    private static readonly Dictionary<string, List<Mod>> ModByPackageIdLookUp = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, List<Mod>> ModByModNameLookUp = new(StringComparer.OrdinalIgnoreCase);
    
    private static List<Mod>? LocalModsCache = null;
    private static List<Mod>? WorkshopModsCache = null;
    private static List<Mod>? OfficialModsCache = null;
    
    public static void ResetCache()
    {
        WorkshopModsCache = null;
        LocalModsCache = null;
        OfficialModsCache = null;
        
        ModByPackageIdLookUp.Clear();
        ModByModNameLookUp.Clear();
        
        foreach (var mod in ModRegistry.Values)
            mod.Dispose();
        ModRegistry.Clear();
    }
    
    /// <summary>공식 컨텐츠의 루트 경로를 반환합니다.</summary>
    public static IEnumerable<string> ModRootsOfficial
    {
        get
        {
            var dirOfficial = Path.Combine(SettingManager.Current.PathRimworld, "Data");
            if (Directory.Exists(dirOfficial))
                foreach (var dir in Directory.EnumerateDirectories(dirOfficial))
                    yield return dir;
        }
    }

    /// <summary>로컬 모드의 루트 경로를 반환합니다.</summary>
    public static IEnumerable<string> ModRootsLocal
    {
        get
        {
            var dirLocalMods = Path.Combine(SettingManager.Current.PathRimworld, "Mods");
            if (Directory.Exists(dirLocalMods))
                foreach (var dir in Directory.EnumerateDirectories(dirLocalMods))
                    yield return dir;
        }
    }

    /// <summary>창작마당 모드의 루트 경로를 반환합니다.</summary>
    public static IEnumerable<string> ModRootsWorkshop
    {
        get
        {
            var dirWorkshopMods = SettingManager.Current.PathWorkshop;
            if (Directory.Exists(dirWorkshopMods))
                foreach (var dir in Directory.EnumerateDirectories(dirWorkshopMods))
                    yield return dir;
        }
    }

    /// <summary>검색된 모든 모드의 루트 경로를 반환합니다.</summary>
    public static IEnumerable<string> ModRootsAll =>
        ModRootsOfficial.Concat(ModRootsLocal).Concat(ModRootsWorkshop);

    public static IEnumerable<Mod> OfficialMods
    {
        get
        {
            if (OfficialModsCache is null)
            {
                OfficialModsCache = ModRootsOfficial.Select(GetModByModRoot)
                    .Where(m => m != null)
                    .OrderBy(x => x!.Name)
                    .ToList()!;
            }
            return OfficialModsCache;
        }
    }
    
    public static IEnumerable<Mod> LocalMods
    {
        get
        {
            if (LocalModsCache == null)
            {
                LocalModsCache = ModRootsLocal.Select(GetModByModRoot).Where(m =>
                    m != null).OrderBy(x => x!.Name).ToList()!;
            }
            return LocalModsCache;
        }
    }

    public static IEnumerable<Mod> WorkshopMods
    {
        get
        {
            if (WorkshopModsCache == null)
            {
                WorkshopModsCache = ModRootsWorkshop.Select(GetModByModRoot).Where(m =>
                    m != null).OrderBy(x => x!.Name).ToList()!;
            }
            return WorkshopModsCache;
        }
    }

    public static IEnumerable<Mod> AllMods => OfficialMods.Concat(LocalMods).Concat(WorkshopMods);
    
    /// <summary>지정한 packageId를 가진 모든 모드 목록을 반환합니다.</summary>
    internal static IReadOnlyList<Mod> GetModsByPackageId(string? packageId)
    {
        if (string.IsNullOrWhiteSpace(packageId)) return Array.Empty<Mod>();

        if (ModByPackageIdLookUp.TryGetValue(packageId, out var cached))
            return cached;

        var matches = AllMods.Where(x => string.Equals(x.PackageID, packageId, StringComparison.OrdinalIgnoreCase)).ToList();
        ModByPackageIdLookUp[packageId] = matches;
        
        return matches;
    }
    
    /// <summary>지정한 모드 이름을 가진 모든 모드 목록을 반환합니다.</summary>
    internal static IReadOnlyList<Mod> GetModsByModName(string? modName)
    {
        if (string.IsNullOrWhiteSpace(modName)) return Array.Empty<Mod>();

        if (ModByModNameLookUp.TryGetValue(modName, out var cached))
            return cached;

        var matches = AllMods.Where(x =>
            string.Equals(x.Name, modName, StringComparison.OrdinalIgnoreCase)).ToList();
        ModByModNameLookUp[modName] = matches;
        
        return matches;
    }
    
    /// <summary>
    /// 모드의 물리적 루트 경로를 기반으로 Mod 객체를 반환합니다. 이미 생성된 객체가 있다면 ModRegistry에서 가져옵니다.
    /// </summary>
    public static Mod? GetModByModRoot(string modRoot)
    {
        // 1. 이미 로드된 모드인지 확인
        string normalizedPath = Path.TrimEndingDirectorySeparator(Path.GetFullPath(modRoot));
        
        if (ModRegistry.TryGetValue(normalizedPath, out var existingMod))
            return existingMod;

        // 2. 새 Mod 객체 생성 (이 과정에서 자동으로 ModRegistry에 추가됨)
        try
        {
            var newMod = new Mod(modRoot, SettingManager.Current);
            ModRegistry[normalizedPath] = newMod;
            return newMod;
        }
        catch (Exception e)
        {
            Log.Err($"[{modRoot}] 모드 로드 실패: {e.Message}");
            return null;
        }
    }
    
    /// <summary>target이 직간접적으로 참조하는 모든 모드를 재귀적으로 찾습니다.</summary>
    /// <remarks>같은 ID를 가지는 다른 모드가 중복으로 존재할 수 있습니다.</remarks>
    public static IEnumerable<Mod> FindAllReferenceMods(Mod target)
    {
        foreach (var officialMod in OfficialMods.Where(official => official.Equals(target)))
        {
            yield return officialMod;
        }

        var set = new HashSet<Mod>();
        ModByPackageIdLookUp.Clear();

        IEnumerable<Mod> FindAllReferenceModsInternal(Mod mod)
        {
            // 1. About.xml에 명시된 종속성 검사
            if (mod.About.modDependencies.Any())
            {
                foreach (var modDependency in mod.About.modDependencies)
                {
                    // 중복된 ID를 가진 모드가 있을 수 있으므로 전부 가져와서 yield 반환
                    var matches = GetModsByPackageId(modDependency);
                    foreach (var match in matches) yield return match;
                }
            }
            
           
            if (Version.TryParse(SettingManager.Current.CurrentVersion, out var targetVersion))
            {
                // About.xml (버전별 종속성)
                if (mod.About.modDependenciesByVersion.TryGetValue(targetVersion, out var versionDeps))
                {
                    foreach (var versionDep in versionDeps)
                    {
                        var matches = GetModsByPackageId(versionDep);
                        foreach (var match in matches) yield return match;
                    }
                }
                
                // 2. LoadFolders.xml 조건부 로드에서 요구하는 패키지 ID 검사
                if (mod.LoadFolders != null && mod.LoadFolders.Rules.TryGetValue(targetVersion, out var rules))
                {
                    var allConditionPids = rules.ifActive.Keys.SelectMany(k => k)
                                            .Concat(rules.ifActiveAll.Keys.SelectMany(k => k))
                                            .Distinct(StringComparer.OrdinalIgnoreCase);

                    foreach (var pid in allConditionPids)
                    {
                        var matches = GetModsByPackageId(pid);
                        foreach (var match in matches) yield return match;
                    }
                }
            }
        }

        IEnumerable<Mod> FindAllReferenceModsRecursive(Mod mod)
        {
            foreach (var child in FindAllReferenceModsInternal(mod))
            {
                if (set.Contains(child)) continue;
                
                yield return child;
                set.Add(child);
                
                foreach (var childchild in FindAllReferenceModsRecursive(child))
                {
                    yield return childchild;
                    set.Add(childchild);
                }
            }
        }

        foreach (var refMod in FindAllReferenceModsRecursive(target))
        {
            yield return refMod;
        }
        ModByPackageIdLookUp.Clear();
    }
}