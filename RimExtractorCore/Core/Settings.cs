using System.Text.Json.Serialization;

namespace RimExtractorCore;

public class Settings
{
    // 1. 기본 경로 및 버전 (기존 Prefabs의 기본값 복원)
    public string PathRimworld { get; set; } = "C:\\Program Files (x86)\\Steam\\steamapps\\common\\RimWorld";
    public string PathWorkshop { get; set; } = "C:\\Program Files (x86)\\Steam\\steamapps\\workshop\\content\\294100";
    //TODO 기본 참조 목록도 필요 없을 것 같아 (애초부터 기본값으로 찍어주는걸로 충분)
    public string PathBaseRefList { get; set; } = "";
    public string CurrentVersion { get; set; } = "1.6";

    // 2. 언어 및 출력 설정
    public string OriginalLanguage { get; set; } = "English";
    public string SecondaryLanguage { get; set; } = string.Empty;
    public string TranslationLanguage { get; set; } = "Korean (한국어)";
    public bool CommentOriginal { get; set; } = false;

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public DuplicateFilePolicy Policy { get; set; } = DuplicateFilePolicy.Overwrite;

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ExportFileFormmat Method { get; set; } = ExportFileFormmat.LanguageData;

    // 비필수 번역 요소(MayNotTranslate) 추출 여부 스위치
    public bool ExtractMayNotTranslate { get; set; } = true;

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

/// <summary>파일 저장시 기존 파일이 있는 경우의 동작을 설정할 때 사용합니다.</summary>
public enum DuplicateFilePolicy { Stop, Overwrite, KeepOriginal }

/// <summary>출력 파일의 포맷을 지정할 때 사용합니다.</summary>
public enum ExportFileFormmat { Spreadsheet, LanguageData, LanguageDataWithComments }