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

        _extractors.Clear();
        foreach (var processor in processors)
        {
            _extractors.Add(processor);
            Log.Msg($"{processor.Name} 등록 완료");
        }

        if (_extractors.Count == 0)
        {
            Log.Err("등록된 IExtractionProcedure가 없습니다.");
        }

        IsInitialized = true;
    }

    // 스냅샷을 받아 등록된 모든 추출기를 순회하며 결과를 합칩니다.
    public IEnumerable<TranslationEntry> Execute(DefSnapshot snapshot, ModMetadata targetMod)
    {
        var combinedEntries = new List<TranslationEntry>();

        foreach (var extractor in _extractors)
        {
            try
            {
                var results = extractor.Extract(snapshot, targetMod);
                if (results != null)
                {
                    combinedEntries.AddRange(results);
                }
            }
            catch (Exception e)
            {
                Log.Err($"추출기({extractor.Name}) 실행 중 오류 발생: {e.Message}");
            }
        }

        return combinedEntries;
    }
}