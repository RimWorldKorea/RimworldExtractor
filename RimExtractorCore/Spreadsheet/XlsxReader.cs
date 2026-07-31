using System.Text;
using ExcelDataReader;

namespace RimExtractorCore.Spreadsheet;

public class XlsxReader : ISpreadsheetReader
{
    static XlsxReader()
    {
        // 파일 인코딩 설정에 따라 읽지 못하게 되는 경우를 예방합니다.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
    
    public bool CanRead(string filePath) => 
        filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);

    public Grid ReadGrid(string filePath, int sheetIndex = 0)
    {
        using var stream = File.OpenRead(filePath);
        // ExcelDataReader는 ClosedXML과 달리 LibreOffice로 편집된 xlsx 파일의 읽기 오류가 없습니다.
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

        // 행 단위 스트리밍 읽기
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