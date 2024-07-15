
static class ExcelComparer
{
    public static void Compare(string path1, string path2)
    {
        var wb1 = Workbook.Load(path1);
        var wb2 = Workbook.Load(path2);
        CompareSheets(wb1, wb2);
    }

    private static void CompareSheets(Workbook wb1, Workbook wb2)
    {
        if (wb1.Worksheets.Count != wb2.Worksheets.Count)
        {
            Console.WriteLine("Number of worksheets don't match.");
        }

        for (int i = 0; i < wb1.Worksheets.Count; i++)
        {
            if (wb1.Worksheets[i].SheetName != wb2.Worksheets[i].SheetName)
            {
                Console.WriteLine($"Sheet name mismatch: {wb1.Worksheets[i].SheetName} - {wb2.Worksheets[i].SheetName}");
            }

            var columnCount = wb1.Worksheets[i].Columns.Count;
            for (int j = 0; j < columnCount; j++)
            {
                // wb1.Worksheets[i].cell

            }
        }
    }
}
