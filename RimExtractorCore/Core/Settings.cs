using System.Text.Json.Serialization;

namespace RimExtractorCore;

public class Settings
{
    // 1. 기본 경로 및 버전 (기존 Prefabs의 기본값 복원)
    public string PathRimworld { get; set; } = "C:\\Program Files (x86)\\Steam\\steamapps\\common\\RimWorld";
    public string PathWorkshop { get; set; } = "C:\\Program Files (x86)\\Steam\\steamapps\\workshop\\content\\294100";
    /// <todo>이제 필요 없지 않을까?</todo>
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

    /// <summary>비필수 번역 요소(MayNotTranslate) 추출 스위치</summary>
    /// <todo>이것도 필요가 있나?</todo>
    public bool ExtractMayNotTranslate { get; set; } = true;
    /// <summary>추출 모드 선택시 자동 참조 기능 작동 여부</summary>
    public bool AutoReferenceLink { get; set; } = true;
    
    // 4. 내부 로직 (직렬화 무시)
    public IEnumerable<string> GetLanguagePriorityList()
    {
        var list = new List<string>();
        if (!string.IsNullOrWhiteSpace(OriginalLanguage)) list.Add(OriginalLanguage);
        if (!string.IsNullOrWhiteSpace(SecondaryLanguage)) list.Add(SecondaryLanguage);
        if (!list.Contains("English")) list.Add("English");
        return list.Distinct();
    }
    
    public List<string> EnabledDefTreeProcedures { get; set; } = new()
    {
        "바닐라 보완 프로시저"
    };
    
    public List<string> EnabledProcedures { get; set; } = new()
    {
        "기본 추출 프로시저"
    };
    
    public List<string> EnabledTranslationEntryProcedures { get; set; } = new();
    
    public Settings Clone()
    {
        return new Settings()
        {
            PathRimworld = this.PathRimworld,
            PathWorkshop = this.PathWorkshop,
            PathBaseRefList = this.PathBaseRefList,
            CurrentVersion = this.CurrentVersion,
            OriginalLanguage = this.OriginalLanguage,
            SecondaryLanguage = this.SecondaryLanguage,
            TranslationLanguage = this.TranslationLanguage,
            CommentOriginal = this.CommentOriginal,
            Policy = this.Policy,
            Method = this.Method,
            ExtractMayNotTranslate = this.ExtractMayNotTranslate,
            EnabledDefTreeProcedures = new List<string>(this.EnabledDefTreeProcedures),
            EnabledProcedures = new List<string>(this.EnabledProcedures),
            EnabledTranslationEntryProcedures = new List<string>(this.EnabledTranslationEntryProcedures),
            AutoReferenceLink = this.AutoReferenceLink
        };
    }
}

/// <summary>파일 저장시 기존 파일이 있는 경우의 동작을 설정할 때 사용합니다.</summary>
public enum DuplicateFilePolicy { Stop, Overwrite, KeepOriginal }

/// <summary>출력 파일의 포맷을 지정할 때 사용합니다.</summary>
public enum ExportFileFormmat { Spreadsheet, LanguageData, LanguageDataWithComments }