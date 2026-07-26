using System;
using System.Collections.Generic;
using System.IO;
using RimExtractorCore.DataTypes;
using RimExtractorCore.Extractor;

namespace RimExtractorCore.Procedures;

public static class TranslationPipelineRunner
{
    private static readonly List<ITranslationProcedure> Processors = new();
    private static bool _isInitialized = false; // 🟢 중복 실행 방지 플래그

    static TranslationPipelineRunner()
    {
        ReloadProcessors();
    }

    public static void ReloadProcessors()
    {
        if (_isInitialized) return;
        
        Processors.Clear();
        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "Translations");

        // 해당 폴더의 ITranslationProcedure 인터페이스 구현체 전부 로드
        var processors = RoslynScriptRunner.LoadProcessorsFromDirectory<ITranslationProcedure>(baseDir);
        foreach (var processor in processors)
        {
            if (!Processors.Any(p => p.Name == processor.Name))
            {
                Processors.Add(processor);
                Log.Msg($"[TranslationPipelineRunner] 번역 후처리 프로시저 등록 완료: {processor.Name}");
            }
        }
        
        _isInitialized = true;
    }

    public static IEnumerable<TranslationEntry> Execute(IEnumerable<TranslationEntry> entries)
    {
        var result = entries;
        foreach (var processor in Processors)
        {
            try
            {
                result = processor.Process(result);
            }
            catch (Exception e)
            {
                Log.Err($"[TranslationPipelineRunner] 프로시저 실행 에러 ({processor.Name}): {e.Message}");
            }
        }
        return result;
    }
}