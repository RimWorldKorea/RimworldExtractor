namespace RimExtractorCore.Spreadsheet;

/// <summary>
/// 림추출기에서 행열 데이터를 다루는 기본 타입입니다.
/// </summary>
public class Grid
{
    public string Name { get; set; } = string.Empty;
    public List<List<string>> Rows { get; set; } = new();

    public int RowCount => Rows.Count;

    public List<string> GetRow(int rowIndex) =>
        rowIndex >= 0 && rowIndex < Rows.Count ? Rows[rowIndex] : new List<string>();
    
    public void AppendRow(IEnumerable<string> row)
    {
        Rows.Add(row.ToList());
    }

    public void AppendRows(IEnumerable<IEnumerable<string>> rows)
    {
        foreach (var row in rows)
        {
            AppendRow(row);
        }
    }
}