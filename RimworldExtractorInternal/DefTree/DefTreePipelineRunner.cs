using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

namespace RimworldExtractorInternal.DefTree;

public static class DefTreePipelineRunner
{
    private static readonly Dictionary<PipelineStage, List<IDefTreePostProcessor>> ProcessorsByStage = new();

    static DefTreePipelineRunner()
    {
        ReloadProcessors();
    }

    /// <summary>
    /// PostProcessors/PostA, PostB, PostC, PostD 폴더의 IDefTreePostProcessor 스크립트들을 수집합니다.
    /// </summary>
    public static void ReloadProcessors()
    {
        ProcessorsByStage.Clear();
        foreach (PipelineStage stage in Enum.GetValues(typeof(PipelineStage)))
        {
            ProcessorsByStage[stage] = new List<IDefTreePostProcessor>();
        }

        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PostProcessors");

        foreach (PipelineStage stage in Enum.GetValues(typeof(PipelineStage)))
        {
            var stageDirName = $"Post{stage.ToString().Replace("Stage", "")}"; // PostA, PostB, PostC, PostD
            var stageFolderPath = Path.Combine(baseDir, stageDirName);

            // 범용 RoslynScriptRunner를 통해 IDefTreePostProcessor 구현체 로드
            var processors = RoslynScriptRunner.LoadProcessorsFromDirectory<IDefTreePostProcessor>(stageFolderPath);
            foreach (var processor in processors)
            {
                ProcessorsByStage[processor.Stage].Add(processor);
            }
        }
    }

    /// <summary>
    /// 특정 파이프라인 단계에 등록된 프로세서들을 순차적으로 실행하여 XDocument 연쇄 변환을 수행합니다.
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
                Log.Err($"[DefTreePipelineRunner] 프로세서 실행 중 예외 발생 ({processor.GetType().Name}): {e.Message}");
            }
        }

        return resultDoc;
    }
}