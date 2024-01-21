static class CsvConverter
{
    public static void Convert(string input, string output)
    {
        var wb = Workbook.Load(input);
        foreach (var sheet in wb.Worksheets)
        {
            Console.WriteLine($"=== {sheet.SheetName} ===");
            Console.WriteLine($"{sheet.Cells.Count}, {sheet.Columns.Count}");
            var rows = new List<string>();
            var columnCount = GetColumnCount(sheet);

            Console.WriteLine($"Column count {columnCount}");

            var columnModulo = 0;
            var sb = new StringBuilder();
            foreach (var cell in sheet.Cells.Keys)
            {
                if (columnModulo % 3 == 0)
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

            var file = File.Create(output);
            using var sw = new StreamWriter(file);
            foreach (var row in rows)
            {
                Console.WriteLine($"{row}");
                sw.WriteLine($"{row}");
            }
        }
    }

    static int GetColumnCount(Worksheet sheet)
    {
        var columnCount = 0;
        foreach (var cell in sheet.Cells.Keys)
        {
            if (cell.Last() == '1')
            {
                columnCount++;
                continue;
            }

            break;
        }

        return columnCount;
    }
}