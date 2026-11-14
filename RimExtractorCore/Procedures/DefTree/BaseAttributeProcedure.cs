using System.Xml.Linq;
using RimExtractorCore.Procedures;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.DefTreeSimulator.Procedures;

/// <summary>
/// 림월드 어셈블리의 번역 태그 등록 상태를 보완하기 위한 프로시저입니다.
/// </summary>
public class BaseAttributeProcedure : IDefTreeProcedure
{
    public string Name => "바닐라 보완 프로시저";

    public void PostPrePile(XDocument prePiledTree)
    {
        
    }
    
    public void PostCreateTree(DefSnapshot snapshot)
    {
        snapshot.BanDef("SoundDef");
        snapshot.BanDef("AbilityGroupDef");
        snapshot.BanDef("FleckDef");
    }
}