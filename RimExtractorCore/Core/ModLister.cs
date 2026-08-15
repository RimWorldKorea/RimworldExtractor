using System.Xml.Linq;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore;

/// <summary>
/// 현재 환경에 설치된 림월드 컨텐츠를 추적하는 도구입니다.
/// </summary>
public static class ModLister
{
    /// <summary>
    /// 공식 컨텐츠의 루트 경로를 반환합니다.
    /// </summary>
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

    /// <summary>
    /// 로컬 모드의 루트 경로를 반환합니다.
    /// </summary>
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

    /// <summary>
    /// 창작마당 모드의 루트 경로를 반환합니다.
    /// </summary>
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

    /// <summary>
    /// 검색된 모든 모드의 루트 경로를 반환합니다.
    /// 근데 왜 이건 쓰는 곳이 없지?
    /// </summary>
    public static IEnumerable<string> ModRootsAll => ModRootsOfficial.Concat(ModRootsLocal).Concat(ModRootsWorkshop);

    public static IEnumerable<ModMetadata> OfficialMods =>
        ModRootsOfficial.Select(GetModMetadataByModRoot).OrderBy(x => x.ModName);

    public static IEnumerable<ModMetadata> LocalMods
    {
        get
        {
            if (LocalModsCache == null)
            {
                LocalModsCache = ModRootsLocal.Select(GetModMetadataByModRoot).OrderBy(x => x.ModName).ToList();
            }

            return LocalModsCache;
        }
    }

    public static IEnumerable<ModMetadata> WorkshopMods
    {
        get
        {
            if (WorkshopModsCache == null)
            {
                WorkshopModsCache = ModRootsWorkshop.Select(GetModMetadataByModRoot).OrderBy(x => x.ModName)
                    .ToList();
            }

            return WorkshopModsCache;
        }
    }

    public static IEnumerable<ModMetadata> AllMods => OfficialMods.Concat(LocalMods).Concat(WorkshopMods);

    public static void ResetCache()
    {
        WorkshopModsCache = null;
        LocalModsCache = null;
    }

    /// <summary>
    /// 주어진 루트 경로로부터 ModMetadata를 생성합니다.
    /// </summary>
    public static ModMetadata GetModMetadataByModRoot(string modRoot)
    {
        var pathAbout = Path.Combine(modRoot, "About", "About.xml");
        string name = "UNKNOWN";
        string packageId = "UNKNOWN";
        var modDependencies = new List<string>();
        if (File.Exists(pathAbout))
        {
            try
            {
                var doc = XDocument.Parse(File.ReadAllText(pathAbout));
                packageId = doc.Root?.Element("packageId")?.Value ?? "UNKNOWN";
                name = doc.Root?.Element("name")?.Value ?? "UNKNOWN";
                if (name == "UNKNOWN")
                {
                    // Official Contents
                    if (doc.Root?.Element("author")?.Value == "Ludeon Studios")
                    {
                        name = Path.GetFileName(modRoot).Trim();
                        return new ModMetadata(modRoot, "Official", name, packageId, true);
                    }
                }

                var modDependenciesNode = doc.Root?.Element("modDependencies");
                if (modDependenciesNode != null)
                {
                    foreach (var childNode in modDependenciesNode.Elements())
                    {
                        var packageIdModDependencies = childNode.Element("packageId");
                        if (packageIdModDependencies != null)
                            modDependencies.Add(packageIdModDependencies.Value);
                    }
                }

                var modDependenciesByVersionNode = doc.Root?.Element("modDependenciesByVersion");
                if (modDependenciesByVersionNode != null)
                {
                    var nodes = modDependenciesByVersionNode.Element("v" + SettingManager.Current.CurrentVersion)
                        ?.Elements();
                    nodes ??= modDependenciesByVersionNode.Elements().LastOrDefault()?.Elements();
                    if (nodes != null)
                    {
                        foreach (var childNode in nodes)
                        {
                            var packageIdModDependencies = childNode.Element("packageId");
                            if (packageIdModDependencies != null)
                                modDependencies.Add(packageIdModDependencies.Value);
                        }
                    }
                }


            }
            catch (Exception e)
            {
                Log.Err($"{pathAbout}에 있는 About.xml 파일을 읽을 수 없었습니다. {e.Message}");
            }
        }

        var pathPublishedFileId = Path.Combine(modRoot, "About", "PublishedFileId.txt");
        var id = "???";
        if (File.Exists(pathPublishedFileId))
        {
            id = File.ReadAllText(pathPublishedFileId).Trim();
        }
        else if (modRoot.Contains("workshop\\content\\294100"))
        {
            id = Path.GetFileName(modRoot);
        }

        modDependencies = modDependencies.Distinct().ToList();
        return new ModMetadata(modRoot, id, name, packageId, false, modDependencies);
    }

    public static List<ExtractableFolder> GetExtractableFolders(ModMetadata modMetadata, string targetVersion)
    {
        var root = modMetadata.RootDir;
        var sets = new HashSet<ExtractableFolder>(new ExtractableFolderComparer());
        var pathLoadFolders = Path.Combine(root, "LoadFolders.xml");

        // 타겟 하위 폴더 목록 (Defs, Patches, Keyed, Strings 등)
        var targetFolders = new List<string> { "Defs", "Patches", "Keyed", "Assemblies" };
        foreach (var lang in SettingManager.Current.GetLanguagePriorityList())
        {
            var shortLang = lang.Split(' ').First();
            targetFolders.Add(Path.Combine("Languages", lang, "Keyed"));
            targetFolders.Add(Path.Combine("Languages", shortLang, "Keyed"));
            targetFolders.Add(Path.Combine("Languages", lang, "Strings"));
            targetFolders.Add(Path.Combine("Languages", shortLang, "Strings"));
        }
        var distinctTargetFolders = targetFolders.Distinct().ToArray();

        IEnumerable<string> GetExtractableFoldersInternal(string path)
        {
            foreach (var folder in distinctTargetFolders)
            {
                var subDir = Path.Combine(path, folder);
                if (Directory.Exists(subDir))
                {
                    yield return Path.GetRelativePath(root, subDir);
                }
            }
        }

        bool loadFoldersXmlParsed = false;

        // 1. LoadFolders.xml 분석 (현재 버전 노드만 탐색)
        if (File.Exists(pathLoadFolders))
        {
            try
            {
                var doc = XDocument.Parse(File.ReadAllText(pathLoadFolders));
                // 예: <v1.5> 노드 탐색
                var versionNode = doc.Root?.Element("v" + targetVersion);

                if (versionNode != null)
                {
                    loadFoldersXmlParsed = true;
                    foreach (var li in versionNode.Elements("li"))
                    {
                        var requiredPackageIds = li.Attribute("IfModActive")?.Value;
                        // 루트("/") 매핑 처리
                        string targetLoadFolder = (li.Value == "/" || li.Value == "\\") ? root : Path.Combine(root, li.Value);

                        foreach (var extractableFolder in GetExtractableFoldersInternal(targetLoadFolder)
                                     .Select(x => new ExtractableFolder(modMetadata, x, requiredPackageIds, targetVersion)
                                         { LoadFolderRoot = targetLoadFolder }))
                        {
                            sets.Add(extractableFolder);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Log.Wrn($"LoadFolders.xml 파싱 오류 ({pathLoadFolders}): {e.Message}");
            }
        }

        // 2. LoadFolders.xml이 없거나 해당 버전 노드가 없는 경우 (Fallback 로직)
        if (!loadFoldersXmlParsed)
        {
            var versionDir = Path.Combine(root, targetVersion);
            
            // 2-A. 폴더 구조형 모드 (예: /1.5, /Common 폴더가 존재하는 경우)
            if (Directory.Exists(versionDir))
            {
                foreach (var extractableFolder in GetExtractableFoldersInternal(versionDir)
                             .Select(x => new ExtractableFolder(modMetadata, x, null, targetVersion)
                                 { LoadFolderRoot = versionDir }))
                {
                    sets.Add(extractableFolder);
                }

                var commonDir = Path.Combine(root, "Common");
                if (Directory.Exists(commonDir))
                {
                    foreach (var extractableFolder in GetExtractableFoldersInternal(commonDir)
                                 .Select(x => new ExtractableFolder(modMetadata, x, null, "Common")
                                     { LoadFolderRoot = commonDir }))
                    {
                        sets.Add(extractableFolder);
                    }
                }
            }
            // 2-B. 단일 구조형 모드 (버전 폴더 없이 루트에 몽땅 들어있는 경우)
            else
            {
                foreach (var extractableFolder in GetExtractableFoldersInternal(root)
                             .Select(x => new ExtractableFolder(modMetadata, x, null, "default")
                                 { LoadFolderRoot = root }))
                {
                    sets.Add(extractableFolder);
                }
            }
        }

        return sets.ToList();
    }

    public static IEnumerable<ModMetadata> FindAllReferenceMods(ModMetadata target)
    {
        foreach (var officialMod in OfficialMods.Where(official => official != target))
        {
            yield return officialMod;
        }

        var set = new HashSet<ModMetadata>();
        ModMetadataByPackageIdLookUp.Clear();

        IEnumerable<ModMetadata> FindAllReferenceModsInternal(ModMetadata modMetadata)
        {
            if (modMetadata.ModDependencies != null)
            {
                foreach (var modDependency in modMetadata.ModDependencies)
                {
                    var b = TryGetModMetadataByPackageId(modDependency, out var possible);
                    if (possible != null)
                        yield return possible;
                }
            }

            foreach (var extractableFolder in GetExtractableFolders(modMetadata, SettingManager.Current.CurrentVersion))
            {
                if (extractableFolder.RequiredPackageId != null)
                {
                    foreach (var modDependency in extractableFolder.RequiredPackageId.Split(','))
                    {
                        var b = TryGetModMetadataByPackageId(modDependency, out var possible);
                        if (possible != null)
                            yield return possible;
                    }
                }
            }
        }

        IEnumerable<ModMetadata> FindAllReferenceModsRecursive(ModMetadata modMetadata)
        {
            foreach (var child in FindAllReferenceModsInternal(modMetadata))
            {
                if (set.Contains(child))
                    continue;
                yield return child;
                set.Add(child);
                foreach (var childchild in FindAllReferenceModsRecursive(child))
                {
                    yield return childchild;
                    set.Add(childchild);
                }
            }
        }

        foreach (var modMetadata in FindAllReferenceModsRecursive(target))
        {
            yield return modMetadata;
        }

        ModMetadataByPackageIdLookUp.Clear();
    }

    internal static bool TryGetModMetadataByPackageId(string? packageId, out ModMetadata? modMetadata)
    {
        if (packageId == null)
        {
            modMetadata = null;
            return false;
        }

        if (ModMetadataByPackageIdLookUp.TryGetValue(packageId, out var value))
        {
            modMetadata = value;
            return value != null;
        }


        var matches = AllMods
            .Where(x => string.Equals(x.PackageId, packageId, StringComparison.CurrentCultureIgnoreCase)).ToList();
        switch (matches.Count)
        {
            case < 1:
                modMetadata = null;
                ModMetadataByPackageIdLookUp[packageId] = null;
                Log.Wrn($"모드 폴더에서 packageId가 {packageId}인 모드를 찾을 수 없었습니다.");
                return false;
            case 1:
                modMetadata = matches[0];
                ModMetadataByPackageIdLookUp[packageId] = modMetadata;
                return true;
            case > 1:
                modMetadata = matches[0];
                ModMetadataByPackageIdLookUp[packageId] = modMetadata;
                Log.Msg($"중복되는 packageId={packageId}, 중복 갯수={matches.Count}.");
                return true;
        }
    }

    internal static ModMetadata? GetModMetadataByModName(string modName)
    {
        if (ModMetadataByModNameLookUp.TryGetValue(modName, out var value))
        {
            return value;
        }

        var result = AllMods.FirstOrDefault(x =>
            string.Equals(x.ModName, modName, StringComparison.CurrentCultureIgnoreCase));
        ModMetadataByModNameLookUp.Add(modName, result);
        return result;
    }

    private static readonly Dictionary<string, ModMetadata?> ModMetadataByPackageIdLookUp = new();
    private static readonly Dictionary<string, ModMetadata?> ModMetadataByModNameLookUp = new();
    private static List<ModMetadata>? LocalModsCache = null;
    private static List<ModMetadata>? WorkshopModsCache = null;
}