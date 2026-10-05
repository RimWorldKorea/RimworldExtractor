using RimExtractorCore.DefTreeSimulator;
using RimExtractorCore.Extractor;

namespace RimExtractorCore;

/// <summary>
/// <para>RimExtractorCore의 Main 역할입니다.</para>
/// </summary>
/// <remarks>실제 프로그램 진입점인 RimExtractorFace의 Initialize 호출에 의해 초기화됩니다.</remarks>

public static class Engine
{
    // 앱 실행 동안 유지될 PrePiledTree XML의 절대 경로
    public static string PrePiledTreePath { get; private set; } = string.Empty;
    
    // 인터페이스를 기반으로 한 Injector 파이프라인 레지스트리
    private static readonly List<IProcedureInjector> Injectors = new()
    {
        DefTreeProcedureInjector.Instance,
        TranslationEntryProcedureInjector.Instance,
        ExtractionProcedureInjector.Instance
    };
    
    public static void Initialize()
    {
        Log.Msg("림추출기 초기화 시작.", true);
        
        // 캐시된 PrePiledTree 버전 체크 및 갱신
        PreparePrePiledTree();
        
        // 인젝터에 프로시저 등록
        foreach (var injector in Injectors)
        {
            injector.RegisterProcedures();
        }
        
        Log.Good("림추출기 초기화 완료.", true);
    }
    
    /// <summary>
    /// 미리 해석된 림월드 DefTree 파일을 로드합니다.
    /// 파일이 없거나 이전 버전이라면 새로 생성합니다.
    /// </summary>
    private static void PreparePrePiledTree()
    {
        var outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PrePiled");
        PrePiledTreePath = Path.Combine(outDir, $"Assembly-{SettingManager.GetFormattedFullVersion()}.xml");

        // 현재 버전에 맞는 XML 파일이 없다면 (버전업이 되었거나 처음 실행인 경우)
        if (!File.Exists(PrePiledTreePath))
        {
            Log.Msg($"신규 버전 감지. ({SettingManager.GetFullVersion()})\n사전 해석 모델 생성 시작.", true);
            
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
            
            var foundFiles = Directory.GetFiles(SettingManager.Current.PathRimworld,
                "Assembly-CSharp.dll", SearchOption.TopDirectoryOnly)
                .ToList();
            
            var subDirs = Directory.GetDirectories(SettingManager.Current.PathRimworld)
                .Where(d => !Path.GetFileName(d).Equals("Mods", StringComparison.OrdinalIgnoreCase));
            
            foreach (var dir in subDirs)
            {
                foundFiles.AddRange(Directory.GetFiles(dir, 
                    "Assembly-CSharp.dll", SearchOption.AllDirectories));
            }
            
            if (foundFiles.Count == 1)
            {
                AssemblyResolver.GenerateBaseDefTree(foundFiles.First(), PrePiledTreePath);
            }
            else if (foundFiles.Count > 1)
            {
                Log.Err("다음 경로에서 여러 개의 Assembly-CSharp.dll을 찾았습니다.", true);
                foreach (var file in foundFiles)
                    Log.Msg($"{file}", true);
                
            }
            else
            {
                Log.Err("림월드 게임 폴더에서 Assembly-CSharp.dll을 찾지 못했습니다.", true);
            }
        }
        else
        {
            Log.Msg($"설치된 림월드 버전과 일치하는 사전 해석 모델 파일 확인. ({SettingManager.GetFullVersion()})", true);
        }
    }
}