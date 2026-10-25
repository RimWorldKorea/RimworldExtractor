namespace RimExtractorCore;

/// <summary>림월드의 LanguageInfo에 대응합니다. 각 언어별 번역 파일에 동봉되어 있습니다.</summary>
public class LanguageInfo : IEquatable<LanguageInfo>
{
    /// <summary>
    /// <para>사용 가능한 모든 언어 목록입니다.</para>
    /// LanguageInfo의 EnglishName을 Key로 사용합니다.
    /// </summary>
    /// <todo>아마 딕셔너리에 넣을 필요는 없을 것 같긴 해.</todo>
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

    /// <summary>언어 이름에 대응하는 LanguageInfo를 반환합니다.</summary>
    public static bool TryParse(string? languageName, out LanguageInfo? result)
    {
        result = null;
        if (string.IsNullOrWhiteSpace(languageName)) return false;
        
        var input = languageName.Trim();
        
        // 입력을 English Name으로 가정
        if (RegisteredLanguages.TryGetValue(input, out result)) return true;

        // 입력을 Native/Extended Name으로 가정
        result = RegisteredLanguages.Values.FirstOrDefault(l => 
            string.Equals(l.NativeName, input, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(l.ExtendedName,  input, StringComparison.OrdinalIgnoreCase));

        if (result is not null) return true;
        
        Log.Wrn($"\"{languageName}\"(은)는 허용된 언어 이름이 아닙니다.", true);
        return false;
    }

    public bool Equals(LanguageInfo? other)
    {
        if (ReferenceEquals(this, other)) return true;
        if (other is null) return false;
        
        return this.ExtendedName == other.ExtendedName ;
    }

    public override bool Equals(object? obj) => Equals(obj as LanguageInfo);

    public override int GetHashCode() => ExtendedName.GetHashCode();
}