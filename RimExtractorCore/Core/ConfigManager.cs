using System.Text.Json;
using System.Text.RegularExpressions;

namespace RimExtractorCore
{
    /// <summary>
    /// 모든 유저 설정은 나에게로
    /// </summary>
    public static class SettingManager
    {
        public const string ConfigFileName = "Settings.json";
        
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
        /// <returns></returns>
        public static string GetFormattedFullVersion()
        {
            //TODO AutoDetectRimworldVersion과 역할이 겹치는 것 같은데
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

        public static string AutoDetectRimworldVersion()
        {
            try
            {
                var pathVersion = Path.Combine(Current.PathRimworld, "Version.txt");
                if (File.Exists(pathVersion))
                {
                    var context = File.ReadAllText(pathVersion).Trim();
                    var match = Regex.Match(context, Current.PatternVersion);
                    if (match.Success)
                    {
                        return match.Groups[0].Value;
                    }
                }
            }
            catch (Exception e)
            {
                Log.Err($"버전 자동 감지 실패: {e.Message}");
            }
            return Current.CurrentVersion;
        }
    }
}