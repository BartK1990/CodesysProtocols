namespace CodesysProtocols.Spreadsheet;

public static class RangeExtensions
{
    public static void Apply(this Range range, Action<int, int> cellAction) => Apply(range, [cellAction]);

    public static void Apply(this Range range, Action<int, int>[] cellActions)
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
