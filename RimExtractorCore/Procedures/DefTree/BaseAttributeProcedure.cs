using System.Xml.Linq;
using RimExtractorCore.Procedures;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.DefTreeSimulator.Procedures;

/// <summary>
/// 림월드 자체 어셈블리의 미흡한 번역 태그 등록 상태를 보완하기 위한 프로시저입니다.
/// </summary>
public class BaseAttributeProcedure : IXDocumentProcedure
{
    public string Name => "기본 ";
    public InjectionStage Stage => InjectionStage.StageA;

    public XDocument Process(XDocument defTree)
    {
        // ---------------------------------------------------------
        // 모드 호환성을 위한 수동 어트리뷰트 주입 영역
        // ---------------------------------------------------------
        
        // 사용 예시: 확장 메서드 덕분에 코드가 매우 짧고 직관적입니다.
        // defTree.TryFindDefAndSetAttribute("AlienRace.ThingDef_AlienRace", "alienProps", Constants.AttrNoTranslate, "True");
        // defTree.TryFindTypeAndSetAttribute("SomeMod.ModSettings", "internalData", Constants.AttrTranslationCanChangeCount, "True");

        return defTree;
    }
}