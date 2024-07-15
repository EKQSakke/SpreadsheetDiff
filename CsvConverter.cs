static class CsvConverter
{
    public static void Convert(string input, string output)
    {
        var wb = Workbook.Load(input);
        var rows = ParseSheet(wb);
        var file = File.Create(output);
        using var sw = new StreamWriter(file);
        foreach (var row in rows)
        {
            sw.WriteLine($"{row}");
        }
    }

    static List<string> ParseSheet(Workbook wb)
    {
        var rows = new List<string>();
        foreach (var sheet in wb.Worksheets)
        {
            rows.Add(sheet.SheetName);
            ParseCells(sheet, rows);
        }
        return rows;
    }

    static void ParseCells(Worksheet sheet, List<string> rows)
    {
        var sb = new StringBuilder();
        var columnCount = GetColumnCount(sheet);
        var columnModulo = 0;
        foreach (var cell in sheet.Cells.Keys)
        {
            if (columnModulo % columnCount == 0)
            {
                if (sb.Length != 0)
                {
                    rows.Add(sb.ToString());
                }

                sb.Clear();
            }
            else
            {
                sb.Append(";");
            }
            sb.Append(sheet.Cells[cell].Value);
            columnModulo++;
        }
        rows.Add(sb.ToString());
    }

    static int GetColumnCount(Worksheet sheet)
    {
        return sheet.Cells.Keys.TakeWhile(x => x.Last() == 1).Count();
    }
}