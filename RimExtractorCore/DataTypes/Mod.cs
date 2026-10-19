using System.Xml.Linq;

namespace RimExtractorCore.DataTypes;

/// <summary>폴더로 구분되는 각각의 모드에 대응합니다.</summary>
public class Mod : IEquatable<Mod>, IDisposable
{
    private static readonly object RegistryLock = new();
    public static HashSet<Mod> RegisteredMods = new();
    
    public readonly string? Path;
    
    public readonly About About;
    public readonly LoadFolders? LoadFolders;

    public readonly HashSet<FolderLoadUnit> OwnedFolders = new();
    
    public string? Name => About.name;
    public string? PackageID => About.packageID;
    public string? Author => About.author;
    public string? WorkshopID => About.publishedFileId;
    
    public readonly ModType Type;
    
    private bool _disposed = false;

    /// <summary>실제 모드 데이터로 모드 참조자를 생성합니다.</summary>
    /// <param name="path">모드 폴더 경로입니다. 유효한 경로가 아니면 예외를 줄거예요.</param>
    /// <param name="extractorSettings">SettingManager에서 현재 설정을 받아 넣으세요. 주로 경로 참조를 위해 필요합니다.</param>
    public Mod(string path, Settings extractorSettings)
    {
        // path는 무조건 유효해야함
        Path = Directory.Exists(path) ? path : throw new DirectoryNotFoundException(path);
        
        About = new About(this);
        
        LoadFolders = File.Exists(System.IO.Path.Combine(path, "LoadFolders.xml")) ?
            new LoadFolders(this) : null;
        
        // 공식 모드야?
        if (Path.IsSubPathOf(System.IO.Path.Combine(extractorSettings.PathRimworld, "Data")))
        {
            Type = ModType.Official;
            
            // 공식 콘텐츠는 About.xml에 모드 이름이 없기 때문에 폴더 이름을 사용
            if (string.IsNullOrWhiteSpace(About.name))
                About.name = System.IO.Path.GetFileName(Path);
        }
        // 창작마당 모드야?
        else if (Path.IsSubPathOf(extractorSettings.PathWorkshop))
        {
            Type = ModType.Workshop;
        }
        // 로컬 모드야?
        else if (Path.IsSubPathOf(System.IO.Path.Combine(extractorSettings.PathRimworld, "Mods")))
        {
            Type = ModType.Local;
        }
        else throw new ArgumentException($"{path} is not a valid mod path.");

        lock (RegistryLock) { RegisteredMods.Add(this); }
    }
    
    /// <summary>실제 모드 데이터 없이 PackageID만으로 가상 모드 참조자를 생성합니다.</summary>
    public Mod(string packageID)
    {
        About = new About(this);
        LoadFolders = new LoadFolders(this);
        
        About.packageID = packageID;
        Type = ModType.Virtual;
        
        lock (RegistryLock) { RegisteredMods.Add(this); }
    }
    
    /// <summary>명시적으로 리소스를 해제하고 레지스트리에서 제거합니다.</summary>
    public void Dispose()
    {
        if (_disposed) return;

        // 자식 FolderLoadUnit들을 먼저 일괄 해제
        foreach (var folder in OwnedFolders.ToList())
        {
            folder.Dispose();
        }
        OwnedFolders.Clear();

        lock (RegistryLock) { RegisteredMods.Remove(this); }
        _disposed = true;
        GC.SuppressFinalize(this);
    }
    
    public override int GetHashCode()
    {
        // 경로가 있으면 경로 자체가 고유값
        if (!string.IsNullOrWhiteSpace(Path))
        {
            string normalizedPath = System.IO.Path.TrimEndingDirectorySeparator(System.IO.Path.GetFullPath(Path));
            return StringComparer.OrdinalIgnoreCase.GetHashCode(normalizedPath);
        }
    
        // 그게 아니면 PackageID가 고유값이 되는 Virtual Mod
        if (!string.IsNullOrWhiteSpace(PackageID))
        {
            return StringComparer.OrdinalIgnoreCase.GetHashCode(PackageID);
        }
    
        // 여기까지 와서는 안 될 것이다.
        Log.Err("유효하지 않은 해시 생성 요청입니다.");
        return base.GetHashCode();
    }
    
    public bool Equals(Mod? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        
        // 둘 다 물리적 경로가 있음
        if (!string.IsNullOrWhiteSpace(this.Path) && !string.IsNullOrWhiteSpace(other.Path))
        {
            return this.Path.IsSamePathWith(other.Path);
        }
    
        // 둘 다 물리적 경로가 없음
        if (string.IsNullOrWhiteSpace(this.Path) && string.IsNullOrWhiteSpace(other.Path))
        {
            if (string.IsNullOrWhiteSpace(this.PackageID) || string.IsNullOrWhiteSpace(other.PackageID))
                return false;
            
            return string.Equals(this.PackageID, other.PackageID, StringComparison.OrdinalIgnoreCase);
        }

        // 하나만 물리적 경로가 존재
        return false;
    }
    
    public override bool Equals(object? obj) => Equals(obj as Mod);
    
    //TODO 이거 꼭 필요한가?
    public string Identifier
    {
        get
        {
            if (Type == ModType.Official) return Name ?? "Unknown";
            var id = string.IsNullOrWhiteSpace(WorkshopID) ? "???" : WorkshopID;
            return $"{Name ?? "Unknown"} - {id}";
        }
    }

    //TODO 이거 꼭 필요한가?
    public override string ToString()
    {
        return $"{(Type == ModType.Official ? "Official" : "Mod")}:{Identifier}";
    }
}

/// <summary>
/// 모드의 About 폴더에 대응합니다.
/// </summary>
public class About
{
    public readonly Mod Parent;
    /// <remarks>About.xml이 아닌 About 폴더의 경로입니다.</remarks>
    public string? Path => string.IsNullOrWhiteSpace(Parent.Path) ? null :
        System.IO.Path.Combine(Parent.Path, "About");
    
    public string? name;
    public string? packageID;
    public string? author;
    public List<string> modDependencies = new ();
    /// <summary>림월드에서 이 설정은 해당 버전에서 modDependencies를 덮어씁니다.</summary>
    public Dictionary<Version, List<string>> modDependenciesByVersion = new ();
    public string? publishedFileId;
    
    public About(Mod parent)
    {
        Parent = parent;
        if (string.IsNullOrEmpty(Path) || !Directory.Exists(Path)) return;
        
        var xmlPath = System.IO.Path.Combine(Path, "About.xml");
        
        try
        {
            if (File.Exists(xmlPath))
            {
                var doc = FileInterface.ReadXml(xmlPath);
                if (doc.Root is null) return;

                packageID = doc.Root.Element("packageId")?.Value ?? null;
                name = doc.Root.Element("name")?.Value ?? null;
                author = doc.Root.Element("author")?.Value ?? null;

                // modDependencies 파싱
                var depNode = doc.Root.Element("modDependencies");
                if (depNode != null)
                {
                    foreach (var li in depNode.Elements("li"))
                    {
                        var pid = li.Element("packageId")?.Value;
                        if (!string.IsNullOrWhiteSpace(pid))
                            modDependencies.Add(pid);
                    }

                    // 중복 의존성 제거
                    modDependencies = modDependencies.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                }

                // modDependenciesByVersion 파싱
                var depByVersionNode = doc.Root.Element("modDependenciesByVersion");
                if (depByVersionNode != null)
                {
                    foreach (var versionNode in depByVersionNode.Elements())
                    {
                        // "v1.4", "v1.5" 등에서 'v' 제거
                        string versionStr = versionNode.Name.LocalName;
                        if (versionStr.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                            versionStr = versionStr.Substring(1);

                        if (Version.TryParse(versionStr, out var parsedVersion))
                        {
                            var versionDeps = new List<string>();
                            foreach (var li in versionNode.Elements("li"))
                            {
                                var pid = li.Element("packageId")?.Value;
                                if (!string.IsNullOrWhiteSpace(pid))
                                    versionDeps.Add(pid);
                            }

                            modDependenciesByVersion[parsedVersion] =
                                versionDeps.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                        }
                    }
                }
            }
        }
        catch (Exception e)
        {
            Log.Err($"About.xml 파싱 실패 ({xmlPath}): {e.Message}");
        }

        // PublishedFileId.txt 확인 (워크샵 모드인 경우)
        var pathPublishedFileId = System.IO.Path.Combine(Path, "PublishedFileId.txt");
        if (File.Exists(pathPublishedFileId))
            publishedFileId = File.ReadAllText(pathPublishedFileId).Trim();
    }

}

/// <summary>
/// 모드의 LoadFolders 파일에 대응합니다.
/// </summary>
public class LoadFolders
{
    public readonly Mod Parent;
    public string Path => System.IO.Path.Combine(Parent.Path!, "LoadFolders.xml");
    
    public readonly Dictionary<Version, LoadFoldersRules> Rules = new ();
    
    public LoadFolders(Mod parent)
    {
        Parent = parent;
        // 파일 존재 여부는 미리 확인하고 생성
        if (!File.Exists(Path)) throw new FileNotFoundException();

        try
        {
            var doc = FileInterface.ReadXml(Path);
            if (doc.Root is null) return;
            
            foreach (var versionNode in doc.Root.Elements())
            {
                string versionStr = versionNode.Name.LocalName;
                
                if (versionStr.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                    versionStr = versionStr.Substring(1);
                
                if (Version.TryParse(versionStr, out Version? parsedVersion))
                    Rules[parsedVersion] = new LoadFoldersRules(versionNode);
            }
        }
        catch (Exception e)
        {
            Log.Wrn($"LoadFolders.xml 파싱 중 오류 발생 ({Path}): {e.Message}");
        }
    }
    
    /// <summary>LoadFolders의 각 필드를 저장하기 위해 사용합니다.</summary>
    /// <remarks>빈 문자열 요소는 모드 루트("/")입니다. 모든 문자열 값은 소문자로 정규화됩니다.</remarks>
    public class LoadFoldersRules
    {
        public readonly HashSet<string> defaultLoad;
        public readonly OrderedDictionary<HashSet<string>, HashSet<string>> ifActive;
        public readonly OrderedDictionary<HashSet<string>, HashSet<string>> ifActiveAll;
        
        /// <param name="loadList">LoadFolders 내의 li 노드 개체입니다.</param>
        public LoadFoldersRules(XElement loadList)
        {
            defaultLoad = new();
            ifActive = new();
            ifActiveAll = new();
            
            foreach (var li in loadList.Elements("li"))
            {
                string folderPath = li.Value.Trim();
                
                // 1. 경로 정규화: 루트 기호는 빈 문자열로, 나머지는 소문자로
                if (folderPath == "/" || folderPath == "\\") 
                    folderPath = "";
                
                folderPath = folderPath.ToLowerInvariant();

                // 2. IfModActive (조건부 로드 - Any)
                var ifModActive = li.Attribute("IfModActive")?.Value;
                if (!string.IsNullOrWhiteSpace(ifModActive))
                {
                    AddStringCondition(ifActive, ifModActive, folderPath);
                    continue;
                }

                // 3. IfModActiveAll (조건부 로드 - All)
                var ifModActiveAll = li.Attribute("IfModActiveAll")?.Value;
                if (!string.IsNullOrWhiteSpace(ifModActiveAll))
                {
                    AddStringCondition(ifActiveAll, ifModActiveAll, folderPath);
                    continue;
                }

                // 4. 조건이 없는 경우 기본 로드
                defaultLoad.Add(folderPath);
            }
        }
        
        /// <summary>문자열을 파싱하여 정규화된 문자열 묶음을 만들고 딕셔너리에 매핑합니다.</summary>
        private void AddStringCondition(OrderedDictionary<HashSet<string>, HashSet<string>> targetDict, 
            string packageIdsCsv, string folderPath)
        {
            // 모든 PackageID를 소문자로 정규화하여 저장
            var pids = packageIdsCsv.Split(',')
                .Select(p => p.Trim().ToLowerInvariant())
                .Where(p => !string.IsNullOrWhiteSpace(p));

            var conditionSet = new HashSet<string>(pids);

            // 동일한 내용물의 HashSet이 이미 키로 존재하는지 검사
            var existingKey = targetDict.Keys.FirstOrDefault(k => k.SetEquals(conditionSet));
            
            if (existingKey != null)
            {
                targetDict[existingKey].Add(folderPath);
            }
            else
            {
                targetDict.Add(conditionSet, new HashSet<string> { folderPath });
            }
        }
    }
}

public enum ModType
{
    /// <summary>림월드 자체 컨텐츠 폴더에 위치한 모드입니다.</summary>
    Official,
    /// <summary>창작마당 폴더에 위치한 모드입니다.</summary>
    Workshop,
    /// <summary>로컬 모드 폴더에 위치한 모드입니다.</summary>
    Local,
    /// <summary>컴퓨터에 존재하지는 않지만, 참조를 통해 존재가 확인되는 모드입니다.</summary>
    Virtual
}