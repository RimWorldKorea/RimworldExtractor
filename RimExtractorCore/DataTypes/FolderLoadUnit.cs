using System.Xml.Linq;

namespace RimExtractorCore.DataTypes;

/// <summary>
/// 림월드에 의해 로드 대상으로 지정되어 루트로 취급되는 폴더에 대응하는 참조 형식입니다.
/// 모든 DirectoryInfo 멤버는 실제 경로의 존재 여부를 보장하도록 설계하세요.
/// <remarks>경로 참조 뿐 아니라 실제 데이터도 여기서 보관하게 해볼까 했는데, 확신이 없어서 놔둠.</remarks>
/// </summary>
public class FolderLoadUnit : IEquatable<FolderLoadUnit>, IDisposable
{
    /// <summary>이 폴더를 직접 호출하는 모드입니다.</summary>
    public readonly Mod Parent;
    /// <summary>Parent와의 상대 경로입니다. 빈 문자열인 경우 폴더 루트입니다.</summary>
    public readonly string RelativePath;
    
    /// <summary>이 폴더의 확인된 경로입니다.</summary>
    public readonly DirectoryInfo? Directory;
    
    public readonly DirectoryInfo? Assemblies;
    public readonly DirectoryInfo? Defs;
    public readonly DirectoryInfo? Languages;
    public readonly DirectoryInfo? Patches;
    public readonly DirectoryInfo? Textures;

    public string FullPath => Directory?.FullName ?? string.Empty;

    private bool _disposed = false;
    
    /// <param name="parent">이 폴더의 null이 아닌 소유자입니다.</param>
    /// <param name="relativePath">parent와의 상대 경로입니다.</param>
    public FolderLoadUnit(Mod parent, string relativePath)
    {
        Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        RelativePath = relativePath;

        if (Parent.Path is null) return;
        
        var path = string.IsNullOrEmpty(relativePath)? Parent.Path : Path.Combine(Parent.Path, relativePath);
        
        Directory = System.IO.Directory.Exists(path) ? new DirectoryInfo(path) : null;
        if (Directory is null) return;
        
        var assembliesPath = Path.Combine(FullPath, "Assemblies");
        var defsPath = Path.Combine(FullPath, "Defs");
        var languagesPath = Path.Combine(FullPath, "Languages");
        var patchesPath = Path.Combine(FullPath, "Patches");
        var texturesPath = Path.Combine(FullPath, "Textures");
        
        Assemblies = System.IO.Directory.Exists(assembliesPath) ? new DirectoryInfo(assembliesPath) : null;
        Defs = System.IO.Directory.Exists(defsPath) ? new DirectoryInfo(defsPath) : null;
        Languages = System.IO.Directory.Exists(languagesPath) ? new DirectoryInfo(languagesPath) : null;
        Patches = System.IO.Directory.Exists(patchesPath) ? new DirectoryInfo(patchesPath) : null;
        Textures = System.IO.Directory.Exists(texturesPath) ? new DirectoryInfo(texturesPath) : null;

        Parent.OwnedFolders.Add(this);
    }
    
    public void Dispose()
    {
        if (_disposed) return;
        Parent.OwnedFolders.Remove(this);
        _disposed = true;
        GC.SuppressFinalize(this);
    }
    
    /// <remarks>Parent와 RelativePath를 해시 소스로 사용합니다.</remarks>
    public override int GetHashCode() => HashCode.Combine(Parent, RelativePath);

    public bool Equals(FolderLoadUnit? other)
    {
        if (other is null) return false;
        
        if (ReferenceEquals(this, other)) return true;
        
        if (this.Directory is null && other.Directory is null) return true;
        
        if (this.Directory is null || other.Directory is null) return false;
        
        return this.FullPath.IsSamePathWith(other.FullPath);
    }

    public override bool Equals(object? obj) => Equals(obj as FolderLoadUnit);

    /// <summary>language에 해당하는 로케일 폴더의 DefInjected 디렉토리를 반환합니다.</summary>
    public DirectoryInfo? GetDefInjected(LanguageInfo language)
    {
        if (GetLocaleFolder(language) is not { } locale) return null;
        var defInjectedPath = Path.Combine(locale.FullName, "DefInjected");
        
        if (System.IO.Directory.Exists(defInjectedPath))
            return new DirectoryInfo(defInjectedPath);
        return null;
    }

    /// <summary>language에 해당하는 로케일 폴더의 Keyed 디렉토리를 반환합니다.</summary>
    public DirectoryInfo? GetKeyedPath(LanguageInfo language)
    {
        if (GetLocaleFolder(language) is not { } locale) return null;
        var keyedPath = Path.Combine(locale.FullName, "Keyed");
        
        if (System.IO.Directory.Exists(keyedPath))
            return new DirectoryInfo(keyedPath);
        return null;
    }

    /// <summary>Languages 폴더 아래의 개별 언어 폴더를 찾습니다.</summary>
    //TODO Locale 폴더 목록 캐싱해두면 좋을 듯
    public DirectoryInfo? GetLocaleFolder(LanguageInfo language)
    {
        if (Directory is null)
        {
            Log.Wrn($"경로가 유효하지 않습니다.");
            return null;
        }
        
        if (Languages is null)
        {
            Log.Wrn($"{FullPath}에 Languages 폴더가 없습니다.");
            return null;
        }
        
        var pathWithEnglishLanguageName =
            Path.Combine(Languages.FullName, language.EnglishName);
        var pathWithExtendedLanguageName =
            Path.Combine(Languages.FullName, language.ExtendedName);
        
        bool isEnglishNameDirectoryExists = System.IO.Directory.Exists(pathWithEnglishLanguageName);
        bool isExtendedNameDirectoryExists = System.IO.Directory.Exists(pathWithExtendedLanguageName);

        // 'Korean (한국어)' 형식을 우선 반환
        if (isExtendedNameDirectoryExists)
        {
            if (isEnglishNameDirectoryExists)
                Log.Wrn($"{Languages}에 {language.ExtendedName}(와)과 {language.EnglishName} 폴더가 중복으로 존재합니다.");
            
            Log.Msg($"{FullPath}에서 발견한 언어 폴더");
            return new DirectoryInfo(pathWithExtendedLanguageName);
        }

        if (isEnglishNameDirectoryExists)
        {
            Log.Msg($"{FullPath}에서 발견한 언어 폴더");
            return new DirectoryInfo(pathWithEnglishLanguageName);
        }
        
        Log.Wrn($"{Languages}에 {language.ExtendedName} 폴더가 없습니다.");
        return null;
    }
}