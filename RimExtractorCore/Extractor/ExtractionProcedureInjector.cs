using RimExtractorCore.DataTypes;

namespace RimExtractorCore.Extractor;

public class ExtractionProcedureInjector : IProcedureInjector
{
    public static ExtractionProcedureInjector Instance { get; } = new();
    public bool IsInitialized { get; private set; } = false;

    // 1. 내부 상태로 캡슐화
    private IExtractionProcedure? _primaryExtractor; 

    private ExtractionProcedureInjector() { }

    public void ReloadProcessors()
    {
        if (IsInitialized) return;

        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "Extractor");
        var processors = RoslynScriptRunner.LoadProcessorsFromDirectory<IExtractionProcedure>(baseDir);
        _primaryExtractor = processors.FirstOrDefault(p => p.Name == "DefaultNodeExtractionProcedure") ?? processors.FirstOrDefault();
        
        if (_primaryExtractor != null)
        {
            Log.Msg($"[ExtractionPipelineRunner] 로드 완료: {_primaryExtractor.Name}");
        }
        else
        {
            Log.Err("[ExtractionPipelineRunner] IExtractionProcedure를 찾을 수 없습니다.");
        }
        
        IsInitialized = true;
    }
    
    public IEnumerable<TranslationEntry>? Execute(SimulationResult simResult)
    {
        if (_primaryExtractor == null)
        {
            throw new InvalidOperationException("등록된 ExtractionProcedure가 존재하지 않아 추출을 진행할 수 없습니다.");
        }
        
        return _primaryExtractor?.Extract(simResult);
    }
}