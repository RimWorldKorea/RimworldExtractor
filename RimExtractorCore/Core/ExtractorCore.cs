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
        
        // 캐시된 PrePiledTree 버전 체크 및 갱신
        PreparePrePiledTree();
        
        // 인젝터에 프로시저 등록
        foreach (var injector in Injectors)
        {
            injector.RegisterProcedures();
        }
        
        Log.Msg("초기화 완료");
    }
    
    /// <summary>
    /// 미리 해석된 림월드 DefTree 파일을 로드합니다.
    /// 파일이 없거나 이전 버전이라면 새로 생성합니다.
    /// </summary>
    private static void PreparePrePiledTree()
    {
        var formattedVersion = SettingManager.GetFormattedFullVersion();
        var outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PrePiled");
        PrePiledTreePath = Path.Combine(outDir, $"Assembly-{formattedVersion}.xml");

        // 현재 버전에 맞는 XML 파일이 없다면 (버전업이 되었거나 처음 실행인 경우)
        if (!File.Exists(PrePiledTreePath))
        {
            Log.Msg($"새로운 림월드 버전 감지 ({formattedVersion}). 사전 해석 모델 생성 시작.");
            
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
            Log.Msg($"현재 림월드 버전({formattedVersion})과 일치하는 사전 해석 모델 파일을 찾았습니다.");
        }
    }
}