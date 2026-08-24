namespace RimExtractorCore.DataTypes;

/// <summary>
/// 림월드에 의해 로드 대상으로 지정되어 루트로 취급되는 폴더입니다.
/// 모든 DirectoryInfo 멤버는 실제 경로의 존재 여부를 보장하도록 설계하세요.
/// </summary>
public class FolderLoadUnit : IEquatable<FolderLoadUnit>
{
    /// <summary>이 폴더를 직접 호출하는 모드입니다.</summary>
    public readonly Mod Parent;
    /// <summary>Parent와의 상대 경로입니다.</summary>
    public readonly string RelativePath;
    
    /// <summary>이 폴더의 확인된 경로입니다.</summary>
    public readonly DirectoryInfo? Directory;
    public readonly DirectoryInfo? Assemblies;
    public readonly DirectoryInfo? Languages;
    /// <summary>이 폴더의 Textures 경로입니다. 이미지를 번역해야 하는 경우가 있기 때문에 필요합니다.</summary>
    public readonly DirectoryInfo? Textures;
    
    /// <param name="parent">이 폴더의 null이 아닌 소유자입니다.</param>
    /// <param name="relativePath">parent와의 상대 경로입니다.</param>
    public FolderLoadUnit(Mod parent, string relativePath)
    {
        Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        RelativePath = relativePath;

        if (Parent.Path is null) return;
        var path = Path.Combine(Parent.Path, relativePath);
        Directory = System.IO.Directory.Exists(path) ? new DirectoryInfo(path) : null;
        if (Directory is null) return;
        
        var assembliesPath = Path.Combine(Directory.FullName, "Assemblies");
        var languagesPath = Path.Combine(Directory.FullName, "Languages");
        var texturesPath = Path.Combine(Directory.FullName, "Textures");
        
        Assemblies = System.IO.Directory.Exists(assembliesPath) ? new DirectoryInfo(assembliesPath) : null;
        Languages = System.IO.Directory.Exists(languagesPath) ? new DirectoryInfo(languagesPath) : null;
        Textures = System.IO.Directory.Exists(texturesPath) ? new DirectoryInfo(texturesPath) : null;
    }
    
    public bool Equals(FolderLoadUnit? other)
    {
        if (this.Directory is null || other?.Directory is null) return false;
        
        //TODO 이게 적절한진 살펴보지 않음
        return this.Directory.Equals(other.Directory);
    }

    public DirectoryInfo? GetDefInjected(LanguageInfo language)
    {
        if (GetLocaleFolder(language) is not { } locale) return null;
        var defInjectedPath = Path.Combine(locale.FullName, "DefInjected");
        
        if (System.IO.Directory.Exists(defInjectedPath))
            return new DirectoryInfo(defInjectedPath);
        return null;
    }

    public DirectoryInfo? GetKeyedPath(LanguageInfo language)
    {
        if (GetLocaleFolder(language) is not { } locale) return null;
        var keyedPath = Path.Combine(locale.FullName, "Keyed");
        
        if (System.IO.Directory.Exists(keyedPath))
            return new DirectoryInfo(keyedPath);
        return null;
    }

    /// <summary>Languages 폴더 아래의 개별 언어 폴더를 찾습니다.</summary>
    private DirectoryInfo? GetLocaleFolder(LanguageInfo language)
    {
        if (Directory is null)
        {
            Log.Wrn($"경로가 유효하지 않습니다.");
            return null;
        }
        
        if (Languages is null)
        {
            Log.Wrn($"{Directory.FullName}에 Languages 폴더가 없습니다.");
            return null;
        }
        
        var pathWithEnglishLanguageName =
            Path.Combine(Languages.FullName, language.GetEnglishName);
        var pathWithExtendedLanguageName =
            Path.Combine(Languages.FullName, language.GetExtendedName);
        
        bool isEnglishNameDirectoryExists = System.IO.Directory.Exists(pathWithEnglishLanguageName);
        bool isExtendedNameDirectoryExists = System.IO.Directory.Exists(pathWithExtendedLanguageName);

        // 'Korean (한국어)' 형식을 우선 반환
        if (isExtendedNameDirectoryExists)
        {
            if (isEnglishNameDirectoryExists)
                Log.Wrn($"{Languages}에 {language.GetExtendedName}(와)과 {language.GetEnglishName} 폴더가 중복으로 존재합니다.");

            return new DirectoryInfo(pathWithExtendedLanguageName);
        }

        if (isEnglishNameDirectoryExists)
        {
            return new DirectoryInfo(pathWithEnglishLanguageName);
        }
        
        Log.Wrn($"{Languages}에 {language.GetExtendedName} 폴더가 없습니다.");
        return null;
    }
}