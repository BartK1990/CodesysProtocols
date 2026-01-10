using CodesysProtocols.Model;
using CodesysProtocols.Model.TableData.Iec608705;
using CodesysProtocols.Spreadsheet.ExcelAccess;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;

namespace CodesysProtocols.Blazor.Services.Codesys23;

public class Iec608705ExcelService : IIec608705ExcelService
{
    public async Task<string[]> GetExcelSheetNamesAsync(Stream stream) =>
        await Task.Run(() => GetExcelSheetNames(stream));

    private static string[] GetExcelSheetNames(Stream stream)
    {
        using ExcelDocument package = Excel.GetExcelDocument(stream);
        return Excel.GetSheetNames(package);
    }

    public async Task<Iec608705Table[]> ReadTablesFromExcelAsync(Stream stream) =>
        await Task.Run(() => ReadTablesFromExcel(stream));

    private static Iec608705Table[] ReadTablesFromExcel(Stream stream)
    {
        using var spreadsheetDoc = SpreadsheetDocument.Open(stream, false);
        using var reader = ExcelDocumentReader.Wrap(spreadsheetDoc, null);
        return Iec608705ExcelWorkbookValidation.SheetNames
            .Select(sheetName =>
            {
                var sheetReader = reader.GetSheet(sheetName);
                return ExcelIec608705Table.Read(sheetReader, sheetName);
            }).ToArray();
    }

    public async Task GetExcelFromTablesAsync(Stream outputStream, Iec608705Table[] tables) =>
        await Task.Run(() => WriteTablesToExcel(outputStream, tables));


    private static void WriteTablesToExcel(Stream outputStream, Iec608705Table[] tables)
    {
        if (tables is null || tables.Length == 0)
        {
            throw new ArgumentException("Tables array cannot be null or empty.", nameof(tables));
        }

        string tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
        try
        {
            using var document = ExcelDocument.Create(tempFile, tables[0].Name);
            var sheet = document.Sheets[0];
            ExcelIec608705Table.Write(tables[0], sheet);
            for (int i = 1; i < tables.Length; i++)
            {
                var worksheet = document.AddWorkSheet(tables[i].Name);
                ExcelIec608705Table.Write(tables[i], worksheet);
            }

            document.Save(outputStream);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }
}