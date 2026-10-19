using System.Text.Json;

namespace RimExtractorCore;

/// <summary>
/// 모든 유저 설정은 나에게로
/// </summary>
public static class SettingManager
{
    public const string ConfigFileName = "Settings.json";

    /// <summary>현재 프로그램에 적용중인 설정입니다.</summary>
    public static Settings Current { get; private set; } = new();

    /// <summary>마지막으로 파일로 저장된 설정입니다. </summary>
    /// <remarks>현재 적용 중인 설정이 아닐 수 있습니다.</remarks>
    public static Settings LastSaved { get; set; } = new();

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public static void Initialize()
    {
        Current = LastSaved.Clone();
    }

    public static void LoadFromFile()
    {
        if (!File.Exists(ConfigFileName))
        {
            InitDefault();
            SaveToDisk();
            return;
        }

        try
        {
            var jsonString = File.ReadAllText(ConfigFileName);
            LastSaved = JsonSerializer.Deserialize<Settings>(jsonString, JsonOptions) ?? new Settings();
        }
        catch (Exception ex)
        {
            Log.Wrn($"설정 파일 로드에 실패하여 기본 설정으로 실행합니다. {ex.Message}");
            InitDefault();
            SaveToDisk();
        }
    }
    
    public static void SaveToDisk()
    {
        try
        {
            var jsonString = JsonSerializer.Serialize(LastSaved, JsonOptions);
            File.WriteAllText(ConfigFileName, jsonString);
        }
        catch (Exception ex)
        {
            Log.Err($"설정 저장 실패: {ex.Message}");
        }
    }

    public static void InitDefault()
    {
        LastSaved = new Settings();
        Current = new Settings(); // 바꿔
    }

    /// <summary>림월드 폴더에서 게임 버전을 읽습니다.</summary>
    /// <returns>"1.7.4977 rev1248" 같이 모든 정보가 남아있는 문자열입니다.</returns>
    public static string GetFullGameVersion()
    {
        var pathVersion = Path.Combine(Current.PathRimworld, "Version.txt");

        try
        {
            if (File.Exists(pathVersion))
                return File.ReadAllText(pathVersion).Trim();
        }
        catch (Exception e)
        {
            Log.Err($"Version.txt 읽기 실패.\n{e.Message}");
        }

        return string.Empty;
    }

    /// <summary>게임 폴더에서 정규화된 게임 버전을 읽습니다.</summary>
    /// <returns>"."과 " "이 모두 "_"로 대체된 문자열입니다.</returns>
    public static string GetFormattedFullGameVersion() => GetFullGameVersion().Replace(".", "_").Replace(" ", "_");

    /// <summary>게임 폴더에서 축약된 게임 버전을 읽습니다.</summary>
    /// <returns>"1.2" 같이 주요 부분만 남긴 문자열입니다.</returns>
    public static string GetGameVersion()
    {
        var rawText = GetFullGameVersion();
        if (rawText == string.Empty) return string.Empty;

        // "1.5.4062 rev824" 같은 형태에서 "1.5.4062" 부분만 추출
        var versionString = rawText.Split(' ')[0];

        // System.Version을 이용해 "Major.Minor"만 반환
        if (Version.TryParse(versionString, out var version))
            return $"{version.Major}.{version.Minor}";

        return string.Empty;
    }

    /// <summary>설치된 림월드의 버전을 읽어들입니다.</summary>
    public static bool UpdateGameVersion()
    {
        var version = GetGameVersion();
        if (version != string.Empty)
        {
            var previousVersion = Current.CurrentVersion;
            Current.CurrentVersion = version;
            Log.Good($"게임 버전 갱신 완료. ({previousVersion} → {version})", true);
            return true;
        }

        Log.Err("게임 버전 갱신 실패.", true);
        return false;
    }
}
