namespace CodesysProtocols.Spreadsheet.OfficeImoExtensions;

public static class ExcelRangeExtensions
{
    public static void Apply(this ExcelRange range, Action<int, int> cellAction) => Apply(range, [cellAction]);

    public static void Apply(this ExcelRange range, Action<int, int>[] cellActions)
    {
        for (int row = range.FirstRow; row <= range.LastRow; row++)
        {
            for (int col = range.FirstCol; col <= range.LastCol; col++)
            {
                foreach (var cellAction in cellActions)
                {
                    cellAction(row, col);
                }
            }
        }
    }
}
