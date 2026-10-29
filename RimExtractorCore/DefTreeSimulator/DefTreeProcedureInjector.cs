using System.Xml.Linq;
using RimExtractorCore.Procedures;

namespace RimExtractorCore.DefTreeSimulator;

/// <summary>
/// DefTreeSimulator에서 동작하는 프로시저 인젝터의 구현체입니다.
/// </summary>
public class DefTreeProcedureInjector : IProcedureInjector
{
    public static DefTreeProcedureInjector Instance { get; } = new();

    private readonly Dictionary<InjectionStage, List<IDefTreeProcedure>> _processorsByStage = new();
    public bool IsInitialized { get; private set; } = false;

    private DefTreeProcedureInjector()
    {
        foreach (InjectionStage stage in Enum.GetValues(typeof(InjectionStage)))
        {
            _processorsByStage[stage] = new List<IDefTreeProcedure>();
        }
    }

    public void RegisterProcedures()
    {
        if (IsInitialized) return;
        
        foreach (var list in _processorsByStage.Values) list.Clear();
        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "DefTree");
        
        var processors = ProcedureLoader.LoadProceduresFromDirectory<IDefTreeProcedure>(baseDir);
        foreach (var processor in processors)
        {
            if (_processorsByStage.TryGetValue(processor.Stage, out var list))
            {
                if (!list.Any(p => p.Name == processor.Name))
                {
                    list.Add(processor);
                    Log.Good($"{processor.Name} 등록 완료. ({processor.Stage})", true);
                }
            }
        }
        
        IsInitialized = true;
    }

    public void ExecuteStage(InjectionStage stage, DefSnapshot currentSnapshot)
    {
        if (!_processorsByStage.TryGetValue(stage, out var processors) ||
            processors.Count == 0)
            return;
        
        foreach (var processor in processors)
        {
            // 화이트리스트 설정에 포함된 경우만 실행
            if (!SettingManager.Current.EnabledDefTreeProcedures.Contains(processor.Name)) continue;
            
            try
            {
                processor.Process(currentSnapshot);
            }
            catch (Exception e)
            {
                Log.Err($"런타임 에러 ({processor.Name}): {e.Message}");
            }
        }
    }
    
}