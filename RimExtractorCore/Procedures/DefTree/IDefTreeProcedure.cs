using System.Xml.Linq;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.Procedures;

/// <summary>
/// DefTreeSimulator에서 실행될 프로시저들의 공통 인터페이스입니다.
/// 이 단계에선 해석 모델 자체의 정확도에 관여하는 코드를 넣어주세요.
/// </summary>
public interface IDefTreeProcedure : IProcedure
{
    /// 개별 프로시저는 PostPrePile와 PostCreateTree 중 필요한 것만 구현하면 됩니다.
    /// 둘 다 필수로 정의할 필요는 없습니다.
    
    /// <summary>사전 해석 트리를 생성한 직후 실행됩니다.</summary>
    void PostPrePile(XDocument prePiledTree) { }
    
    /// <summary> 상속(XmlInheritance) 처리가 끝난 후 추출 직전에 실행됩니다. </summary>
    void PostCreateTree(DefSnapshot snapshot) { }
}