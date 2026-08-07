using RimExtractorCore.DataTypes;
using RimExtractorCore.Exceptions;

namespace RimExtractorCore.Spreadsheet;

public static class SpreadsheetReader
{
    private static readonly List<ISpreadsheetReader> Readers = new()
    {
        new XlsxReader(),
        new OdsReader()
    };

    /// <summary>
    /// 포맷에 맞는 리더를 찾아 2차원 데이터(Grid)를 읽어옵니다.
    /// </summary>
    public static Grid ReadGrid(string filePath, int sheetIndex = 0)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"파일을 찾을 수 없습니다: {filePath}");

        var reader = Readers.Find(r => r.CanRead(filePath))
            ?? throw new NotSupportedException($"지원하지 않는 포맷입니다: {Path.GetExtension(filePath)}");

        return reader.ReadGrid(filePath, sheetIndex);
    }

    /// <summary>
    /// 파일에서 Grid를 읽어와 TranslationEntry 리스트로 변환합니다.
    /// </summary>
    public static List<TranslationEntry> ReadTranslations(string filePath)
    {
        var grid = ReadGrid(filePath);
        return ParseTranslations(grid);
    }

    /// <summary>
    /// Grid 데이터를 해석하여 TranslationEntry 리스트를 생성합니다.
    /// </summary>
    public static List<TranslationEntry> ParseTranslations(Grid grid)
    {
        var translations = new List<TranslationEntry>();
        if (grid.RowCount == 0) return translations;

        var headerRow = grid.Rows.First();

        int colClass = FindColumnIndex(headerRow, ISpreadsheetReader.HeaderClass);
        if (colClass == -1) throw new XlsxHeaderReadingException("Class");

        int colNode = FindColumnIndex(headerRow, ISpreadsheetReader.HeaderNode);
        if (colNode == -1) throw new XlsxHeaderReadingException("Node");

        int colRequiredMods = FindColumnIndex(headerRow, ISpreadsheetReader.HeaderRequiredMods);
        
        int colMayNotNecessary = FindColumnIndex(headerRow, ISpreadsheetReader.HeaderMayNotNecessary);

        int colOriginal = FindColumnIndex(headerRow, ISpreadsheetReader.HeaderOriginal);
        if (colOriginal == -1) throw new XlsxHeaderReadingException("Original");

        int colTranslated = FindColumnIndex(headerRow, ISpreadsheetReader.HeaderTranslated);
        if (colTranslated == -1) throw new XlsxHeaderReadingException("Translated");

        for (int i = 1; i < grid.Rows.Count; i++)
        {
            var row = grid.Rows[i];

            string className = GetValSafely(row, colClass);
            string node = GetValSafely(row, colNode);

            if (string.IsNullOrEmpty(className) || string.IsNullOrEmpty(node)) continue;

            RequiredMods? requiredMods = null;
            string textRequiredMods = GetValSafely(row, colRequiredMods);
            if (!string.IsNullOrEmpty(textRequiredMods))
            {
                if (textRequiredMods.Contains('\n'))
                {
                    requiredMods = new RequiredMods();
                    foreach (var s in textRequiredMods.Split('\n'))
                    {
                        requiredMods.AddAllowedByModName(s);
                    }
                }
                else
                {
                    requiredMods = RequiredMods.FromStringByModNames(textRequiredMods);
                }
            }
            
            bool mayNotNecessary = false;
            if (colMayNotNecessary != -1)
            {
                // [수정됨] 엑셀에서 읽어 들일 때도 하드코딩 대신 상수로 비교
                mayNotNecessary = GetValSafely(row, colMayNotNecessary).Contains(Constants.AttrMayNotTranslate, StringComparison.OrdinalIgnoreCase);
            }

            string original = GetValSafely(row, colOriginal);
            string rawTranslated = GetValSafely(row, colTranslated);
            string? translated = string.IsNullOrEmpty(rawTranslated) ? null : rawTranslated;

            //TODO 뭔가 이상한데
            translations.Add(new TranslationEntry(className, node, original, translated, requiredMods, null) { MayNotNecessary = mayNotNecessary });
        }

        return translations;
    }

    private static int FindColumnIndex(List<string> headerRow, IEnumerable<string> candidateNames)
    {
        foreach (var candidate in candidateNames)
        {
            int index = headerRow.FindIndex(h => string.Equals(h.Trim(), candidate, StringComparison.OrdinalIgnoreCase));
            if (index != -1) return index;
        }
        return -1;
    }

    private static string GetValSafely(List<string> row, int index)
    {
        if (index >= 0 && index < row.Count) return row[index];
        return string.Empty;
    }
}