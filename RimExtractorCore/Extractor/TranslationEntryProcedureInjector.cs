using RimExtractorCore.DataTypes;
using RimExtractorCore.Procedures;

namespace RimExtractorCore.Extractor;

/// <summary>
/// Extractor 동작 끝 부분에서 작동하는 프로시저의 인젝터 구현체입니다.
/// </summary>
public class TranslationEntryProcedureInjector : IProcedureInjector
{
    public static TranslationEntryProcedureInjector Instance { get; } = new();

    private readonly List<ITranslationEntryProcedure> _processors = new();
    public bool IsInitialized { get; private set; } = false;

    private TranslationEntryProcedureInjector() { }

    public void RegisterProcedures()
    {
        if (IsInitialized) return;
        
        _processors.Clear();
        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "Translations");
        
        var processors = ProcedureLoader.LoadProceduresFromDirectory<ITranslationEntryProcedure>(baseDir);
        foreach (var processor in processors)
        {
            if (!_processors.Any(p => p.Name == processor.Name))
            {
                _processors.Add(processor);
                Log.Good($"{processor.Name} 등록.", true);
            }
        }
        
        IsInitialized = true;
    }

    public IEnumerable<TranslationEntry> Execute(IEnumerable<TranslationEntry> entries)
    {
        var result = entries;
        foreach (var processor in _processors)
        {
            // 화이트리스트 설정에 포함된 경우만 실행
            if (!SettingManager.Current.EnabledTranslationEntryProcedures.Contains(processor.Name)) continue;
            
            try
            {
                result = processor.Process(result);
            }
            catch (Exception e)
            {
                Log.Err($"런타임 에러 ({processor.Name}): {e.Message}");
            }
        }
        return result;
    }
}