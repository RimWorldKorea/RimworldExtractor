using System.Xml.Linq;

namespace RimExtractorCore.DefTree;

public enum PipelineStage
{
    StageA, // 어셈블리 베이스 사전 구성 이후 (사전 참조 모델 구성 전)
    StageB, // 사전 참조 모델 구성 이후 (패치 오퍼레이션 적용 전)
    StageC, // 패치 오퍼레이션 적용 이후 (랭귀지 데이터 오버라이드 전)
    StageD  // 랭귀지 데이터 오버라이드 이후 (최종 DefTree 반환 전)
}

public interface IXDocumentProcedure
{
    /// <summary>
    /// 프로시저의 고유 식별자 이름입니다.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 프로시저가 실행될 파이프라인 지점입니다.
    /// </summary>
    PipelineStage Stage { get; }

    /// <summary>
    /// XDocument(DefTree)를 전달받아 수정 후 다음 프로시저 또는 단계로 전달합니다.
    /// </summary>
    XDocument Process(XDocument defTree);
}