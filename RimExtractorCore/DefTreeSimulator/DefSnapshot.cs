using System.Xml.Linq;
using RimExtractorCore.DataTypes;

public class DefSnapshot
{
    public List<string> RequiredModIds { get; set; } = new();
    public XDocument Tree { get; set; }

    // [NEW] 이 스냅샷(우주)에서 로드되어야 할 폴더들
    public List<ExtractableFolder> AssignedFolders { get; set; } = new();
    
    // [NEW] 이 스냅샷(우주)에서 실행되어야 할 패치 폴더들
    public List<ExtractableFolder> AssignedPatches { get; set; } = new();

    public DefSnapshot(XDocument tree, IEnumerable<string>? requiredModIds = null)
    {
        Tree = tree;
        if (requiredModIds != null) RequiredModIds.AddRange(requiredModIds);
    }

    public DefSnapshot Clone(IEnumerable<string> additionalModIds)
    {
        var clonedTree = new XDocument(Tree);
        var newIds = new List<string>(RequiredModIds);
        newIds.AddRange(additionalModIds);
        
        return new DefSnapshot(clonedTree, newIds)
        {
            // 복제될 때 폴더 할당 상태도 그대로 가져갑니다.
            AssignedFolders = new List<ExtractableFolder>(this.AssignedFolders),
            AssignedPatches = new List<ExtractableFolder>(this.AssignedPatches)
        };
    }
}