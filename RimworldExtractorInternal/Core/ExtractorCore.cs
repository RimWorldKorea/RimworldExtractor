using System.Threading.Tasks;
using RimworldExtractorInternal.DefTree;
using RimworldExtractorInternal.Procedures;

namespace RimworldExtractorInternal.Core;

public static class ExtractorCore
{
    /// <summary>
    /// 동적 프로시저(.cs)를 Roslyn으로 컴파일하고 파이프라인에 등록합니다.
    /// </summary>
    public static void Initialize()
    {
        Log.Msg("[ExtractorCore] 프로시저 초기화 및 동적 컴파일 시작...");
        
        DefTreePipelineRunner.ReloadProcessors();
        TranslationPipelineRunner.ReloadProcessors();
        
        Log.Msg("[ExtractorCore] 프로시저 초기화 완료.");
    }
}