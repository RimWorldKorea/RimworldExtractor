using RimExtractorCore.DataTypes;

namespace RimExtractorCore.Procedures;

/// <summary>
/// DefTreeSimulator의 시뮬레이션 모델로부터 번역 데이터를 추출하기 위한 프로시저의 인터페이스입니다.
/// 이 타입의 프로시저는 단 하나만 존재할 수 있습니다.
/// </summary>
public interface IExtractionProcedure
{
    string Name { get; }
    IEnumerable<TranslationEntry> Extract(SimulationResult simResult);
}