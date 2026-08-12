using System.Text.Json.Serialization;

namespace RimExtractorCore;

public class Settings
{
    // 1. 기본 경로 및 버전 (기존 Prefabs의 기본값 복원)
    public string PathRimworld { get; set; } = "C:\\Program Files (x86)\\Steam\\steamapps\\common\\RimWorld";
    public string PathWorkshop { get; set; } = "C:\\Program Files (x86)\\Steam\\steamapps\\workshop\\content\\294100";
    public string PathBaseRefList { get; set; } = "";
    public string CurrentVersion { get; set; } = "1.6";

    // 2. 언어 및 출력 설정
    public string OriginalLanguage { get; set; } = "English";
    public string SecondaryLanguage { get; set; } = string.Empty;
    public string TranslationLanguage { get; set; } = "Korean (한국어)";
    public bool CommentOriginal { get; set; } = false;

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public DuplicatesPolicy Policy { get; set; } = DuplicatesPolicy.Overwrite;

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ExtractionMethod Method { get; set; } = ExtractionMethod.Languages;

    // 비필수 번역 요소(MayNotTranslate) 추출 여부 스위치
    public bool ExtractMayNotNecessary { get; set; } = true;

    // 4. 내부 로직 (직렬화 무시)

    public IEnumerable<string> GetLanguagePriorityList()
    {
        var list = new List<string>();
        if (!string.IsNullOrWhiteSpace(OriginalLanguage)) list.Add(OriginalLanguage);
        if (!string.IsNullOrWhiteSpace(SecondaryLanguage)) list.Add(SecondaryLanguage);
        if (!list.Contains("English")) list.Add("English");
        return list.Distinct();
    }
}

//TODO enum들은 왜 따로 빠져있지?

/// 저장하려는 곳에 중복 파일이 있는 경우 행동 지정
public enum DuplicatesPolicy { Stop = 0, Overwrite, KeepOriginal }
/// 출력 파일의 형식 지정
public enum ExtractionMethod { Excel = 0, Languages, LanguagesWithComments }