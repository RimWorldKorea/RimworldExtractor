using ExcelDataReader;

namespace RimworldExtractorInternal.Spreadsheet;

public class XlsxReader : ISpreadsheetReader
{
    public bool CanRead(string filePath) => 
        filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);

    public Grid ReadGrid(string filePath, int sheetIndex = 0)
    {
        // .NET 환경에서 한글/특수 인코딩 호환성을 확보하기 위한 등록
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using var stream = File.OpenRead(filePath);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        // 타깃 시트 위치로 이동
        int currentSheet = 0;
        while (currentSheet < sheetIndex && reader.NextResult())
        {
            currentSheet++;
        }

        var grid = new Grid 
        { 
            Name = reader.Name ?? $"Sheet{sheetIndex + 1}" 
        };

        // 행 단위 스트리밍 읽기 (LibreOffice 주석/스타일 버그 완전 무시)
        while (reader.Read())
        {
            var rowData = new List<string>();
            int fieldCount = reader.FieldCount;

            for (int col = 0; col < fieldCount; col++)
            {
                var val = reader.GetValue(col);
                rowData.Add(val?.ToString() ?? string.Empty);
            }

            grid.Rows.Add(rowData);
        }

        return grid;
    }
}