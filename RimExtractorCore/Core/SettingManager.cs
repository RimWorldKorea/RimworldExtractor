using System.Text.Json;

namespace RimExtractorCore
{
    /// <summary>
    /// 모든 유저 설정은 나에게로
    /// </summary>
    public static class SettingManager
    {
        public const string ConfigFileName = "Settings.json";
        
        /// <summary>
        /// 현재 적용된 설정은 전부 여기서 조회합니다.
        /// </summary>
        public static Settings Current { get; private set; } = new();

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions 
        { 
            WriteIndented = true 
        };

        public static void Load()
        {
            if (!File.Exists(ConfigFileName))
            {
                InitDefault();
                Save();
                return;
            }

            try
            {
                var jsonString = File.ReadAllText(ConfigFileName);
                Current = JsonSerializer.Deserialize<Settings>(jsonString, JsonOptions) ?? new Settings();
            }
            catch (Exception ex)
            {
                Log.Wrn($"설정 파일 로드에 실패하여 기본 설정으로 실행합니다. {ex.Message}");
                InitDefault();
                Save();
            }
        }

        public static void Save()
        {
            try
            {
                var jsonString = JsonSerializer.Serialize(Current, JsonOptions);
                File.WriteAllText(ConfigFileName, jsonString);
            }
            catch (Exception ex)
            {
                Log.Err($"설정 저장 실패: {ex.Message}");
            }
        }

        public static void InitDefault()
        {
            Current = new Settings();
        }
        
        /// <summary>
        /// 림월드 폴더에서 게임 버전 기록을 읽어 특정 형태로 가공합니다.
        /// </summary>
        public static string GetFormattedFullVersion()
        {
            try
            {
                var pathVersion = Path.Combine(Current.PathRimworld, "Version.txt");
                if (File.Exists(pathVersion))
                {
                    var rawContext = File.ReadAllText(pathVersion).Trim();
                    // 점(.)과 공백( )을 언더바(_)로 치환
                    return rawContext.Replace(".", "_").Replace(" ", "_");
                }
            }
            catch (Exception e)
            {
                Log.Err($"버전 파일 읽기 실패: {e.Message}");
            }
            return "UnknownVersion";
        }

        /// <summary>
        /// 설치된 림월드의 버전을 읽어들입니다.
        /// </summary>
        public static string AutoDetectRimworldVersion()
        {
            try
            {
                var pathVersion = Path.Combine(Current.PathRimworld, "Version.txt");
                
                if (File.Exists(pathVersion))
                {
                    var context = File.ReadAllText(pathVersion).Trim();
                    
                    // "1.5.4062 rev824" 같은 형태에서 "1.5.4062" 부분만 추출
                    var versionString = context.Split(' ')[0];

                    // System.Version을 이용해 안전하게 파싱 (성공 시 Major.Minor만 반환)
                    if (Version.TryParse(versionString, out var version))
                    {
                        return $"{version.Major}.{version.Minor}";
                    }
                }
            }
            catch (Exception e)
            {
                Log.Err($"버전 감지 실패: {e.Message}");
            }
            
            return Current.CurrentVersion;
        }
    }
}