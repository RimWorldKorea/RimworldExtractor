namespace RimExtractorCore.DataTypes;

/// <summary>
/// 폴더로 구분되는 각각의 모드에 대응합니다.
/// </summary>
public class Mod
{
    public readonly string? Path;
    
    public readonly About About;
    public readonly LoadFolders LoadFolders;
    
    public string? Name => About.name;
    public string? PackageID => About.packageID;
    public string? Author => About.author;
    public string? WorkshopID => About.publishedField;
    
    public readonly ModType Type;

    /// <summary>
    /// 실제 모드 데이터로 모드 참조자를 생성합니다.
    /// </summary>
    public Mod(string path, Settings extractorSettings)
    {
        About = new About(this);
        LoadFolders = new LoadFolders(this);
        
        Path = path;
        // extractorSettings로부터 Type 계산
    }

    /// <summary>
    /// 실제 모드 데이터 없이 PackageID만으로 가상 모드 참조자를 생성합니다.
    /// </summary>
    public Mod(string packageID)
    {
        About = new About(this);
        LoadFolders = new LoadFolders(this);
        
        About.packageID = packageID;
        Type = ModType.Virtual;
    }
}

/// <summary>
/// 모드의 About 폴더에 대응합니다.
/// </summary>
public class About
{
    public readonly Mod Parent;
    public string? Path => Parent.Path is null ? null : System.IO.Path.Combine(Parent.Path, "About");
    
    public string? name;
    public string? packageID;
    public string? author;
    public List<string> modDependencies = new List<string>();
    public string? publishedField;
    
    public About(Mod parent)
    {
        Parent = parent;
        // 모드의 About 폴더 경로일 aboutDirectory 내의 파일을 읽어서 필드를 채움
    }

}

/// <summary>
/// 모드의 LoadFolders 파일에 대응합니다.
/// </summary>
public class LoadFolders
{
    public readonly Mod Parent;
    public string? Path => Parent.Path is null ? null : System.IO.Path.Combine(Parent.Path, "LoadFolders.xml");

    public readonly HashSet<FolderLoadUnit> defaultLoad;
    public readonly OrderedDictionary<HashSet<Mod>, HashSet<FolderLoadUnit>> ifActive;
    public readonly OrderedDictionary<HashSet<Mod>, HashSet<FolderLoadUnit>> ifActiveAll;
    
    public LoadFolders(Mod parent)
    {
        defaultLoad = new HashSet<FolderLoadUnit>();
        ifActive = new OrderedDictionary<HashSet<Mod>, HashSet<FolderLoadUnit>>();
        ifActiveAll = new OrderedDictionary<HashSet<Mod>, HashSet<FolderLoadUnit>>();
        
        Parent = parent;
        
        // 모드의 LoadFolders.xml 경로일 loadFoldersPath를 읽어서 필드를 채움
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