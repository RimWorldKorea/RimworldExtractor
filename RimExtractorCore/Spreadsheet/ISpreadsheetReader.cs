namespace RimExtractorCore.Spreadsheet;

public interface ISpreadsheetReader
{
    // 공통 헤더 매칭 후보 목록 (우선순위 순서)
    public static string[] HeaderClass => new[] { FileInterface.HeaderClass, "Class [Not chosen]"/*하위 호환용*/ };
    public static string[] HeaderNode => new[] { FileInterface.HeaderIdentifier, "Node [Not chosen]"/*하위 호환용*/ };
    public static string[] HeaderRequiredMods => new[] { FileInterface.HeaderRequiredMods, "Required Mods [Not chosen]"/*하위 호환용*/ };
    public static string[] HeaderMayNotNecessary => new[] { FileInterface.HeaderMayNotNecessary };
    
    public static string[] HeaderOriginal => new[] 
    { 
        $"{SettingManager.Current.OriginalLanguage} [Source string]"/*하위 호환용*/, 
        "EN [Source string]"/*하위 호환용*/, 
        FileInterface.HeaderOriginal 
    };
    
    public static string[] HeaderTranslated => new[] 
    { 
        $"{SettingManager.Current.TranslationLanguage} [Translation]"/*하위 호환용*/, 
        "KO [Translation]"/*하위 호환용*/, 
        FileInterface.HeaderTranslation 
    };

    bool CanRead(string filePath);

    /// <summary>
    /// 파일 경로를 받아 2차원 순수 데이터(Grid)를 반환합니다.
    /// </summary>
    Grid ReadGrid(string filePath, int sheetIndex = 0);
}