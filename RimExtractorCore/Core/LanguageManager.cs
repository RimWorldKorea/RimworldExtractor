namespace RimExtractorCore;

/// <summary>림월드의 LanguageInfo에 대응합니다. 각 언어별 번역 파일에 동봉되어 있습니다.</summary>
public class LanguageInfo
{
    private readonly string friendlyNameNative;
    private readonly string friendlyNameEnglish;

    LanguageInfo(string friendlyNameEnglish, string friendlyNameNative)
    {
        this.friendlyNameEnglish = friendlyNameEnglish;
        this.friendlyNameNative = friendlyNameNative;
    }
    
    public string GetNativeName => friendlyNameNative;
    public string GetEnglishName => friendlyNameEnglish;
    public string GetExtendedName => $"{friendlyNameEnglish} ({friendlyNameNative})";
}