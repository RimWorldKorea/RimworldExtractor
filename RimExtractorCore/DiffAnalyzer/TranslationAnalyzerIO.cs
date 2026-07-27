using RimExtractorCore.DataTypes;

namespace RimExtractorCore.DiffAnalyzer;

public static class TranslationAnalyzerIO
{
    /// <summary>
    /// [WIP] 기존 IO.ModifyExcel 역할.
    /// 추후 ClosedXML 대신 ODS 포맷(SpreadsheetReader/Writer)을 활용하여 
    /// 메모리 상에서 Grid를 병합하고 덮어쓰는 로직으로 재구현할 예정입니다.
    /// </summary>
    public static void ModifySpreadsheet(List<TranslationAnalyzerEntry.ChangeRecord> changes, string targetPath)
    {
        throw new NotImplementedException("ODS 기반의 번역 파일 병합 로직으로 재구현이 필요합니다.");
    }
}