namespace RimExtractorCore;

/// <summary>림월드의 LanguageInfo에 대응합니다. 각 언어별 번역 파일에 동봉되어 있습니다.</summary>
public class LanguageInfo
{
    private static readonly Dictionary<string, LanguageInfo> RegisteredLanguages =
        new(StringComparer.OrdinalIgnoreCase);
    
    public static readonly IReadOnlyList<LanguageInfo> SupportedLanguages;
    
    public string NativeName { get; }
    public string EnglishName { get; }
    public string ExtendedName => $"{EnglishName} ({NativeName})";
    
    static LanguageInfo()
    {
        var initializers = new[]
        {
            ("English", "English"),
            ("Catalan", "Català"),
            ("ChineseSimplified", "简体中文"),
            ("ChineseTraditional", "繁體中文"),
            ("Czech", "Čeština"),
            ("Danish", "Dansk"),
            ("Dutch", "Nederlands"),
            ("Estonian", "Eesti"),
            ("Finnish", "Suomi"),
            ("French", "Français"),
            ("German", "Deutsch"),
            ("Greek", "Ελληνικά"),
            ("Hungarian", "Magyar"),
            ("Italian", "Italiano"),
            ("Japanese", "日本語"),
            ("Korean", "한국어"),
            ("Norwegian", "Norsk Bokmål"),
            ("Polish", "Polski"),
            ("Portuguese", "Português"),
            ("PortugueseBrazilian", "Português Brasileiro"),
            ("Romanian", "Română"),
            ("Russian", "Русский"),
            ("Slovak", "Slovenčina"),
            ("Spanish", "Español(Castellano)"),
            ("SpanishLatin", "Español(Latinoamérica)"),
            ("Swedish", "Svenska"),
            ("Turkish", "Türkçe"),
            ("Ukrainian", "Українська"),
            ("Vietnamese", "tiếng Việt"),
        };
        
        var list = new List<LanguageInfo>();
        foreach (var (englishName, nativeName) in initializers)
        {
            var languageInfo = new LanguageInfo(englishName, nativeName);
            RegisteredLanguages[englishName] = languageInfo;
            list.Add(languageInfo);
        }
        
        SupportedLanguages = list.AsReadOnly();
    }
    
    private LanguageInfo(string englishName, string nativeName)
    {
        EnglishName = englishName;
        NativeName = nativeName;
    }

    /// <summary>
    /// 영어 이름을 기반으로 등록된 언어 정보를 반환합니다.
    /// (목록에 없는 커스텀 언어일 경우 임시 객체를 반환합니다)
    /// </summary>
    public static LanguageInfo Get(string englishName)
    {
        if (RegisteredLanguages.TryGetValue(englishName, out var info))
        {
            return info;
        }
        
        Log.Wrn($"[{englishName}] 하드코딩된 언어 목록에 없습니다. 임시 객체를 생성합니다.");
        return new LanguageInfo(englishName, englishName);
    }
}