using OfficeOpenXml;

namespace CodesysProtocols.Spreadsheet.ExcelAccess;

public class Excel
{
    public static ExcelPackage GetExcelPackage(string path)
    {
        FileInfo fileInfo = new(path);
        return new ExcelPackage(fileInfo);
    }

    public static ExcelPackage GetExcelPackage(Stream stream) => new(stream);

    public static ExcelWorksheet GetSheet(ExcelPackage package, string sheetName)
    {
        return package.Workbook.Worksheets[sheetName];
    }

    public static string[] GetSheetNames(ExcelPackage package)
    {
        return package.Workbook.Worksheets.Select(w => w.Name).ToArray();
    }

    public static string? GetCellValue(object cell)
    {
        return cell?.ToString();
    }

    public static int GetLastSheetRow(ExcelWorksheet worksheet)
    {
        if (worksheet.Dimension is null)
        {
            return 0;
        }
        var row = worksheet.Dimension.Rows;
        while (row >= 1)
        {
            var range = worksheet.Cells[row, 1, row, worksheet.Dimension.Columns];
            if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
            {
                break;
            }
            row--;
        }
        return row;
    }

    public static int GetLastSheetColumn(ExcelWorksheet worksheet)
    {
        if (worksheet.Dimension is null)
        {
            return 0;
        }
        var column = worksheet.Dimension.Columns;
        while (column >= 1)
        {
            var range = worksheet.Cells[1, column, worksheet.Dimension.Rows, column];
            if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
            {
                break;
            }
            column--;
        }
        return column;
    }
}