using System.Xml.Linq;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore.DefTreeSimulator;
/// <summary>
/// 다중 로드 분기 환경에서 개별 로드셋의 DefTree 구성을 위한 패키지입니다.
/// </summary>
public class DefSnapshot
{
    public List<string> RequiredModIds { get; set; } = new();
    public XDocument Tree { get; set; }
    
    public List<FolderLoadUnit> LoadedUnits { get; set; } = new();
    
    /// <summary>true인 경우 이 스냅샷이 다른 스냅샷들의 공통 이미지임을 나타냅니다.</summary>
    public bool IsBaseSnapshot { get; set; } = false;

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
            LoadedUnits = new List<FolderLoadUnit>(this.LoadedUnits),
            IsBaseSnapshot = false
        };
    }
}