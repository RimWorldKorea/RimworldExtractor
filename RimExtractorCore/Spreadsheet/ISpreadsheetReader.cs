namespace RimExtractorCore.Spreadsheet;

public interface ISpreadsheetReader
{
    // 🟢 공통 헤더 매칭 후보 목록 (우선순위 순서)
    public static string[] HeaderClass => new[] { "Class", "Class [Not chosen]" };
    public static string[] HeaderNode => new[] { "Node", "Node [Not chosen]" };
    public static string[] HeaderRequiredMods => new[] { "Required Mods", "Required Mods [Not chosen]" };
    
    public static string[] HeaderOriginal => new[] 
    { 
        $"{SettingManager.Current.OriginalLanguage} [Source string]", 
        "EN [Source string]", 
        "Original" 
    };
    
    public static string[] HeaderTranslated => new[] 
    { 
        $"{SettingManager.Current.TranslationLanguage} [Translation]", 
        "KO [Translation]", 
        "Translated" 
    };

    bool CanRead(string filePath);

    /// <summary>
    /// 파일 경로를 받아 2차원 순수 데이터(Grid)를 반환합니다.
    /// </summary>
    Grid ReadGrid(string filePath, int sheetIndex = 0);
}