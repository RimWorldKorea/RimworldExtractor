using RimExtractorCore.DefTreeSimulator;
using RimExtractorCore.Extractor;

namespace RimExtractorCore;

public static class ExtractorCore
{
    // 앱 실행 동안 유지될 PrePiledTree XML의 절대 경로
    public static string PrePiledTreePath { get; private set; } = string.Empty;
    
    // 인터페이스를 기반으로 한 Injector 파이프라인 레지스트리
    private static readonly List<IProcedureInjector> Injectors = new()
    {
        XDocumentProcedureInjector.Instance,
        TranslationEntryProcedureInjector.Instance,
        ExtractionProcedureInjector.Instance
    };

    public static void Initialize()
    {
        Log.Msg("초기화 시작");
        
        // [추가됨] 1. 사전 트리(PrePiledTree) 버전 체크 및 갱신
        PreparePrePiledTree();
        
        // 프로시저 등록
        foreach (var injector in Injectors)
        {
            injector.RegisterProcedures();
        }
        
        Log.Msg("초기화 완료");
    }
    
    private static void PreparePrePiledTree()
    {
        var formattedVersion = SettingManager.GetFormattedFullVersion();
        var outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PrePiled");
        PrePiledTreePath = Path.Combine(outDir, $"Assembly-{formattedVersion}.xml");

        // 현재 버전에 맞는 XML 파일이 없다면 (버전업이 되었거나 처음 실행인 경우)
        if (!File.Exists(PrePiledTreePath))
        {
            Log.Msg($"새로운 림월드 버전 감지됨({formattedVersion}). 사전 트리 생성을 시작합니다.");
            
            if (Directory.Exists(outDir))
            {
                // 이전 버전의 찌꺼기 파일(Assembly-*.xml)들을 싹 정리합니다.
                foreach (var file in Directory.GetFiles(outDir, "Assembly-*.xml"))
                {
                    File.Delete(file);
                }
            }
            else
            {
                Directory.CreateDirectory(outDir);
            }

            var assemblyPath = Path.Combine(SettingManager.Current.PathRimworld, "RimWorldWin64_Data", "Managed", "Assembly-CSharp.dll");
            
            // AssemblyResolver에 저장할 절대 경로를 직접 넘겨줍니다.
            AssemblyResolver.GenerateBaseDefTree(assemblyPath, PrePiledTreePath);
        }
        else
        {
            Log.Msg($"현재 버전({formattedVersion})과 일치하는 캐시된 사전 트리를 발견했습니다.");
        }
    }
}