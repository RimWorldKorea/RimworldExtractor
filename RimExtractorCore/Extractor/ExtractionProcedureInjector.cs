namespace RimExtractorCore.Extractor;

public class ExtractionProcedureInjector : IProcedureInjector
{
    public static ExtractionProcedureInjector Instance { get; } = new();

    public bool IsInitialized { get; private set; } = false;

    private ExtractionProcedureInjector() { }

    public void ReloadProcessors()
    {
        if (IsInitialized) return;

        // DefaultNodeExtractionProcedure.cs를 Procedures/Extractors 폴더에 위치시킵니다.
        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "Extractor");
        
        var processors = RoslynScriptRunner.LoadProcessorsFromDirectory<IExtractionProcedure>(baseDir);
        var primary = processors.FirstOrDefault(p => p.Name == "DefaultNodeExtractionProcedure") ?? processors.FirstOrDefault();
        
        if (primary != null)
        {
            ExtractorEngine.PrimaryExtractor = primary;
            Log.Msg($"[ExtractionPipelineRunner] 로드 완료: {primary.Name}");
        }
        else
        {
            Log.Wrn("[ExtractionPipelineRunner] IExtractionProcedure를 찾을 수 없습니다.");
        }
        
        IsInitialized = true;
    }
}