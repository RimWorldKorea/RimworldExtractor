using RimExtractorCore.DataTypes;
using RimExtractorCore.Procedures;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.Extractor;

/// <summary>
/// Extractor에서 동작하는 프로시저 인젝터의 구현체입니다.
/// Extractor 프로시저는 단 하나만 허용됩니다.
/// </summary>
public class ExtractionProcedureInjector : IProcedureInjector
{
    public static ExtractionProcedureInjector Instance { get; } = new();
    public bool IsInitialized { get; private set; } = false;

    // 단일 Extractor가 아닌 리스트로 관리합니다.
    private readonly List<IExtractionProcedure> _extractors = new();

    private ExtractionProcedureInjector()
    {
    }

    public void RegisterProcedures()
    {

        
        if (IsInitialized) return;

        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "Extractor");
        var processors = ProcedureLoader.LoadProceduresFromDirectory<IExtractionProcedure>(baseDir);

        //todo 이거 1개만 등록시켜야 되는데 그렇게 돼있는게 아닌 것 같아
        
        _extractors.Clear();
        foreach (var processor in processors)
        {
            _extractors.Add(processor);
            Log.Good($"{processor.Name} 등록.", true);
        }

        if (_extractors.Count == 0)
        {
            Log.Err("등록된 IExtractionProcedure가 없습니다.", true);
        }

        IsInitialized = true;
    }

    // 스냅샷을 받아 등록된 모든 추출기를 순회하며 결과를 합칩니다.
    public IEnumerable<TranslationEntry> Execute(DefSnapshot snapshot, Mod targetMod)
    {
        var combinedEntries = new List<TranslationEntry>();

        var activeExtractor =
            _extractors.FirstOrDefault(e => SettingManager.Current.EnabledProcedures.Contains(e.Name));
        
        if (activeExtractor != null)
        {
            try
            {
                var results = activeExtractor.Extract(snapshot, targetMod);
                combinedEntries.AddRange(results);
            }
            catch (Exception e)
            {
                Log.Err($"추출 프로시저 실패 ({activeExtractor.Name}) : {e.Message}");
            }
        }
        else
        {
            Log.Err("활성화된 추출 프로시저가 없습니다.", true);
        }

        return combinedEntries;
    }
}