using RimExtractorCore.DataTypes;

namespace RimExtractorCore.Extractor;

public class TranslationEntryProcedureInjector : IProcedureInjector
{
    public static TranslationEntryProcedureInjector Instance { get; } = new();

    private readonly List<ITranslationEntryProcedure> _processors = new();
    public bool IsInitialized { get; private set; } = false;

    private TranslationEntryProcedureInjector() { }

    public void ReloadProcessors()
    {
        if (IsInitialized) return;
        
        _processors.Clear();
        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "Translations");
        
        var processors = RoslynScriptRunner.LoadProcessorsFromDirectory<ITranslationEntryProcedure>(baseDir);
        foreach (var processor in processors)
        {
            if (!_processors.Any(p => p.Name == processor.Name))
            {
                _processors.Add(processor);
                Log.Msg($"[TranslationPipelineRunner] 로드 완료: {processor.Name}");
            }
        }
        
        IsInitialized = true;
    }

    public IEnumerable<TranslationEntry> Execute(IEnumerable<TranslationEntry> entries)
    {
        var result = entries;
        foreach (var processor in _processors)
        {
            try
            {
                result = processor.Process(result);
            }
            catch (Exception e)
            {
                Log.Err($"[TranslationPipelineRunner] 런타임 에러 ({processor.Name}): {e.Message}");
            }
        }
        return result;
    }
}