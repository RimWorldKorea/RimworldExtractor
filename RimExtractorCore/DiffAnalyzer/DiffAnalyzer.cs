using RimExtractorCore.DataTypes;

namespace RimExtractorCore.DiffAnalyzer;

//TODO 분석기는 내가 안쓰기도 해서 잘 모르겠다.

public static class AnalyzerEngine
{
    private static readonly Dictionary<string, Mod?> ModsByFilePath = new();
    
    public static string[] GetXlsxPaths(string rootPath) =>
        FileInterface.DescendantFiles(rootPath)
            .Where(x => x.ToLower().EndsWith(".xlsx") || x.ToLower().EndsWith(".ods"))
            .ToArray();

    public static Mod? GetModFromFilePath(string filePath)
    {
        if (ModsByFilePath.TryGetValue(filePath, out var value)) return value;
        
        var tmp = GetModFromFilePathInternal(filePath);
        ModsByFilePath[filePath] = tmp;
        return tmp;
    }
    
    /// 우선순위 판별 헬퍼
    private static Mod? SelectBestMod(IReadOnlyList<Mod> mods)
    {
        if (mods.Count == 0) return null;
        if (mods.Count == 1) return mods[0];

        // 중복될 경우 Local(작업 폴더) > Workshop(구독 폴더) > Official 순으로 덮어쓰기 적용
        return mods.OrderBy(m => m.Type switch
        {
            ModType.Local => 1,
            ModType.Workshop => 2,
            ModType.Official => 3,
            _ => 4
        }).First();
    }

    private static Mod? GetModFromFilePathInternal(string filePath)
    {
        var fileName = Path.GetFileNameWithoutExtension(filePath);
        
        // Step 1: (ModName - Id)
        var dashIdx = fileName.LastIndexOf('-');
        if (dashIdx > 0)
        {
            var modName = fileName[..dashIdx].Trim();
            
            var mod = SelectBestMod(ModLister.GetModsByModName(modName.StripInvaildChars()));
            if (mod != null) return mod;

            var id = fileName[(dashIdx + 1)..].Trim();
            mod = ModLister.AllMods.FirstOrDefault(x => x.WorkshopID == id);
            if (mod != null) return mod;
        }

        // Step 2 (ModName)
        {
            var modName = fileName;
            var modMetadata = ModLister.AllMods.FirstOrDefault(x =>
                string.Equals(x.Name.StripInvaildChars(), modName, StringComparison.CurrentCultureIgnoreCase));
            if (modMetadata != null)
            {
                return modMetadata;
            }
        }
        // Step 3 (Id)
        if (long.TryParse(fileName, out _))
        {
            var modMetadata = ModLister.AllMods.FirstOrDefault(x => x.PackageID == fileName);
            if (modMetadata != null)
            {
                return modMetadata;
            }
        }

        // Step 4 (FolderName)
        var dirName = Path.GetFileName(Path.GetDirectoryName(filePath));
        if (dirName == null)
        {
            return null;
        }

        // Step 4-1: (ModName - Id)
        dashIdx = dirName.LastIndexOf('-');
        if (dashIdx > 0)
        {
            var modName = dirName[..dashIdx];
            var modMetadata = ModLister.AllMods.FirstOrDefault(x =>
                string.Equals(x.Name.StripInvaildChars(), modName, StringComparison.CurrentCultureIgnoreCase));
            if (modMetadata != null)
            {
                return modMetadata;
            }

            var id = dirName[(dashIdx + 1)..].Trim();
            modMetadata = ModLister.AllMods.FirstOrDefault(x => x.PackageID == id);
            if (modMetadata != null)
            {
                return modMetadata;
            }
        }

        // Step 4-2 (ModName)
        {
            var modName = dirName;
            var modMetadata = ModLister.AllMods.FirstOrDefault(x =>
                string.Equals(x.Name.StripInvaildChars(), modName, StringComparison.CurrentCultureIgnoreCase));
            if (modMetadata != null)
            {
                return modMetadata;
            }
        }
        // Step 3 (Id)
        if (long.TryParse(dirName, out _))
        {
            var modMetadata = ModLister.AllMods.FirstOrDefault(x => x.PackageID == dirName);
            if (modMetadata != null)
            {
                return modMetadata;
            }
        }

        return null;
    }
}