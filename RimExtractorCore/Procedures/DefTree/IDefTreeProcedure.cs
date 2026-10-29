using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.Procedures;

/// <summary>
/// DefTreeSimulator에서 실행될 프로시저들의 공통 인터페이스입니다.
/// 이 단계에선 추출을 고려하지 않고 시뮬레이션 모델 자체의 정확도를 높이기 위한 규칙을 넣어주세요.
/// </summary>
public interface IDefTreeProcedure : IProcedure
{
    /// <summary>인젝터가 관리하는 프로시저가 실행될 지점입니다.</summary>
    InjectionStage Stage { get; }

    /// <summary>XDocument를 전달받아 가공 후 다음 프로시저 또는 단계로 전달합니다.</summary>
    void Process(DefSnapshot defTree);
}

public enum InjectionStage
{
    PostTreeGeneration, // DefTree 구축 완료 후
}