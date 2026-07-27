using System.Xml.Linq;
using RimExtractorCore.Procedures;

namespace RimExtractorCore.DefTreeSimulator;

public class XDocumentProcedureInjector : IProcedureInjector
{
    public static XDocumentProcedureInjector Instance { get; } = new();

    private readonly Dictionary<PipelineStage, List<IXDocumentProcedure>> _processorsByStage = new();
    public bool IsInitialized { get; private set; } = false;

    private XDocumentProcedureInjector()
    {
        foreach (PipelineStage stage in Enum.GetValues(typeof(PipelineStage)))
        {
            _processorsByStage[stage] = new List<IXDocumentProcedure>();
        }
    }

    public void RegisterProcedures()
    {
        if (IsInitialized) return;
        
        foreach (var list in _processorsByStage.Values) list.Clear();
        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "DefTree");
        
        var processors = ProcedureLoader.LoadProceduresFromDirectory<IXDocumentProcedure>(baseDir);
        foreach (var processor in processors)
        {
            if (_processorsByStage.TryGetValue(processor.Stage, out var list))
            {
                if (!list.Any(p => p.Name == processor.Name))
                {
                    list.Add(processor);
                    Log.Msg($"{processor.Stage}에 {processor.Name} 등록");
                }
            }
        }
        
        IsInitialized = true;
    }

    public XDocument ExecuteStage(PipelineStage stage, XDocument currentDefTree)
    {
        if (!_processorsByStage.TryGetValue(stage, out var processors) || processors.Count == 0)
        {
            return currentDefTree;
        }

        var resultDoc = currentDefTree;
        foreach (var processor in processors)
        {
            try
            {
                resultDoc = processor.Process(resultDoc);
            }
            catch (Exception e)
            {
                Log.Err($"런타임 에러 ({processor.Name}): {e.Message}");
            }
        }
        return resultDoc;
    }
}