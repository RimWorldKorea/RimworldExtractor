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

    // [NEW] 이 스냅샷(우주)에서 로드되어야 할 폴더들
    public List<ExtractableFolder> AssignedFolders { get; set; } = new();
    
    // [NEW] 이 스냅샷(우주)에서 실행되어야 할 패치 폴더들
    public List<ExtractableFolder> AssignedPatches { get; set; } = new();
    
    // [추가] 베이스 스냅샷 여부를 명시적으로 추적하는 플래그
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
            // 복제될 때 폴더 할당 상태도 그대로 가져갑니다.
            AssignedFolders = new List<ExtractableFolder>(this.AssignedFolders),
            AssignedPatches = new List<ExtractableFolder>(this.AssignedPatches),
            
            // 복제(파생)된 스냅샷은 조건이 추가된 Branch이므로 무조건 false
            IsBaseSnapshot = false
        };
    }
}