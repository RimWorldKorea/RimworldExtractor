using RimExtractorCore.DefTreeSimulator;
using RimExtractorCore.Extractor;

namespace RimExtractorCore;

public static class ExtractorCore
{
    // 인터페이스를 기반으로 한 Injector 파이프라인 레지스트리
    private static readonly List<IProcedureInjector> Injectors = new()
    {
        XDocumentProcedureInjector.Instance,
        TranslationEntryProcedureInjector.Instance,
        ExtractionProcedureInjector.Instance
    };

    public static void Initialize()
    {
        Log.Msg("[ExtractorCore] 모듈 초기화를 시작합니다...");
        
        // 공통 인터페이스를 통해 일괄 초기화 수행
        foreach (var injector in Injectors)
        {
            injector.ReloadProcessors();
        }
        
        Log.Msg("[ExtractorCore] 준비 완료.");
    }
}