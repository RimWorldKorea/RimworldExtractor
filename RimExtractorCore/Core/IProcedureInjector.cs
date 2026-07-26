namespace RimExtractorCore;

/// <summary>
/// 프로시저 로딩 파이프라인의 생명주기를 관리하는 공통 인터페이스입니다.
/// </summary>
public interface IProcedureInjector
{
    bool IsInitialized { get; }
    void ReloadProcessors();
}