namespace RimExtractorCore;

/// <summary>
/// 프로시저를 실제로 가져다 쓰는 인젝터 클래스의 공통 인터페이스입니다.
/// </summary>
public interface IProcedureInjector
{
    bool IsInitialized { get; }
    
    void RegisterProcedures();
}