using System.IO.Compression;
using System.Xml.Linq;

namespace RimExtractorCore.Spreadsheet;

public class OdsReader : ISpreadsheetReader
{
    private static readonly XNamespace TableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private static readonly XNamespace TextNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    public bool CanRead(string filePath) => 
        filePath.EndsWith(".ods", StringComparison.OrdinalIgnoreCase);

    public Grid ReadGrid(string filePath, int sheetIndex = 0)
    {
        using var zip = ZipFile.OpenRead(filePath);
        var contentEntry = zip.GetEntry("content.xml")
            ?? throw new FileNotFoundException("ODS 파일 내부에서 content.xml을 찾을 수 없습니다.");

        using var stream = contentEntry.Open();
        var doc = XDocument.Load(stream);

        var table = doc.Descendants(TableNs + "table").ElementAtOrDefault(sheetIndex)
            ?? throw new IndexOutOfRangeException($"시트 인덱스({sheetIndex})를 찾을 수 없습니다.");

        var grid = new Grid
        {
            Name = table.Attribute(TableNs + "name")?.Value ?? $"Sheet{sheetIndex + 1}"
        };

        foreach (var rowElem in table.Elements(TableNs + "table-row"))
        {
            var rowValues = new List<string>();

            foreach (var cellElem in rowElem.Elements(TableNs + "table-cell"))
            {
                int repeatCount = 1;
                var repeatAttr = cellElem.Attribute(TableNs + "number-columns-repeated");
                if (repeatAttr != null && int.TryParse(repeatAttr.Value, out int parsedRepeat))
                {
                    repeatCount = parsedRepeat;
                }

                var cellText = string.Join("\n", cellElem.Elements(TextNs + "p").Select(p => p.Value));

                for (int i = 0; i < repeatCount; i++)
                {
                    rowValues.Add(cellText);
                }
            }

            if (rowValues.Any(v => !string.IsNullOrWhiteSpace(v)))
            {
                grid.Rows.Add(rowValues);
            }
        }

        return grid;
    }
}