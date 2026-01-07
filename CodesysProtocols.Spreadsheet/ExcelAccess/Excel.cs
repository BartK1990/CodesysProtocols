using OfficeIMO.Excel;

namespace CodesysProtocols.Spreadsheet.ExcelAccess;

public class Excel
{
    public static ExcelDocument GetExcelPackage(string path)
    {
        return ExcelDocument.Load(path, true, false);
    }

    public static ExcelDocument GetExcelPackage(Stream stream)
    {
        return ExcelDocument.Load(stream, true, false);
    }

    public static ExcelSheet GetSheet(ExcelDocument package, string sheetName)
    {
        return package.Sheets.First(s => s.Name == sheetName);
    }

    public static string[] GetSheetNames(ExcelDocument package)
    {
        return package.Sheets.Select(w => w.Name).ToArray();
    }

    public static string? GetCellValue(object? cell)
    {
        return cell?.ToString();
    }

    public static int GetLastSheetRow(ExcelSheet worksheet)
    {
        var usedRange = worksheet.UsedRangeA1;
        if (string.IsNullOrEmpty(usedRange))
        {
            return 0;
        }

        // Parse the used range (e.g., "A1:C10") to get the last row
        var parts = usedRange.Split(':');
        if (parts.Length != 2)
        {
            return 0;
        }

        var (lastRow, _) = A1.ParseCellRef(parts[1]);
        return lastRow;
    }

    public static int GetLastSheetColumn(ExcelSheet worksheet)
    {
        var usedRange = worksheet.UsedRangeA1;
        if (string.IsNullOrEmpty(usedRange))
        {
            return 0;
        }

        // Parse the used range (e.g., "A1:C10") to get the last column
        var parts = usedRange.Split(':');
        if (parts.Length != 2)
        {
            return 0;
        }

        var (_, lastColumn) = A1.ParseCellRef(parts[1]);
        return lastColumn;
    }
}