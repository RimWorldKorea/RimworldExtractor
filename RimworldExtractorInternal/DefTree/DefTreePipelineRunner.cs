using System;
using System.Collections.Generic;
using System.IO;
using RimworldExtractorInternal.Core;
using System.Xml.Linq;

namespace RimworldExtractorInternal.DefTree;

public static class DefTreePipelineRunner
{
    private static readonly Dictionary<PipelineStage, List<IXDocumentProcedure>> ProcessorsByStage = new();
    private static bool _isInitialized = false;

    static DefTreePipelineRunner()
    {
        ReloadProcessors();
    }

    /// <summary>
    /// Procedures/ 폴더 내의 모든 외부 .cs 프로시저 파일들을 읽어와 Stage별로 등록합니다.
    /// </summary>
    public static void ReloadProcessors()
    {
        if (_isInitialized) return;
        
        ProcessorsByStage.Clear();
        foreach (PipelineStage stage in Enum.GetValues(typeof(PipelineStage)))
        {
            ProcessorsByStage[stage] = new List<IXDocumentProcedure>();
        }

        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "DefTree");
        
        // Procedures 폴더 내의 모든 .cs 파일에서 IXDocumentProcedure 구현체 로드
        var processors = RoslynScriptRunner.LoadProcessorsFromDirectory<IXDocumentProcedure>(baseDir);
        foreach (var processor in processors)
        {
            if (ProcessorsByStage.TryGetValue(processor.Stage, out var list))
            {
                // 중복 인스턴스 등록 방지 체크
                if (!list.Any(p => p.Name == processor.Name))
                {
                    list.Add(processor);
                    Log.Msg($"[DefTreePipelineRunner] 프로시저 등록 완료: {processor.Name} ({processor.Stage})");
                }
            }
        }
        
        _isInitialized = true;
    }

    /// <summary>
    /// 지정된 Stage의 프로시저들을 순차적으로 적용하여 가공된 XDocument를 반환합니다.
    /// </summary>
    public static XDocument ExecuteStage(PipelineStage stage, XDocument currentDefTree)
    {
        if (!ProcessorsByStage.TryGetValue(stage, out var processors) || processors.Count == 0)
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
                Log.Err($"[DefTreePipelineRunner] 프로시저 실행 에러 ({processor.Name}): {e.Message}");
            }
        }
        return resultDoc;
    }
}