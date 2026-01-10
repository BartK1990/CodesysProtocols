namespace CodesysProtocols.Spreadsheet;

public static class RangeExtensions
{
    public static void Apply(this Range range, Action<int, int> cellAction)
    {
        for (int row = range.FirstRow; row <= range.LastRow; row++)
        {
            for (int col = range.FirstCol; col <= range.LastCol; col++)
            {
                cellAction(row, col);
            }
        }
    }
}
