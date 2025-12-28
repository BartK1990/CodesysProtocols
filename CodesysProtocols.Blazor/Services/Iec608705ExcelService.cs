using CodesysProtocols.Model;
using CodesysProtocols.Model.TableData.Iec608705;
using CodesysProtocols.Spreadsheet.ExcelAccess;
using OfficeOpenXml;

namespace CodesysProtocols.Blazor.Services;

public class Iec608705ExcelService : IIec608705ExcelService
{
    public async Task<string[]> GetExcelSheetNamesAsync(Stream stream) =>
        await Task.Run(() => GetExcelSheetNames(stream));

    private static string[] GetExcelSheetNames(Stream stream)
    {
        using ExcelPackage package = Excel.GetExcelPackage(stream);
        return Excel.GetSheetNames(package);
    }

    public async Task<Iec608705Table[]> ReadTablesFromExcelAsync(Stream stream) =>
        await Task.Run(() => ReadTablesFromExcel(stream));

    private static Iec608705Table[] ReadTablesFromExcel(Stream stream)
    {
        using ExcelPackage package = Excel.GetExcelPackage(stream);
        return Iec608705ExcelWorkbookValidation.SheetNames
            .Select(sheetName => Excel.GetSheet(package, sheetName))
            .Select(ExcelIec608705Table.Read).ToArray();
    }

    public async Task GetExcelFromTablesAsync(Stream outputStream, Iec608705Table[] tables) =>
        await Task.Run(() => WriteTablesToExcel(outputStream, tables));


    private static void WriteTablesToExcel(Stream outputStream, Iec608705Table[] tables)
    {
        using ExcelPackage package = new();
        foreach (Iec608705Table table in tables)
        {
            var worksheet = package.Workbook.Worksheets.Add(table.Name);
            ExcelIec608705Table.Write(table, worksheet);
        }

        package.SaveAs(outputStream);
    }
}