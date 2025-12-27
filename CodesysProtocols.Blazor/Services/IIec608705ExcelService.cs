using CodesysProtocols.Model.TableData.Iec608705;

namespace CodesysProtocols.Blazor.Services;

public interface IIec608705ExcelService
{
    Task<string[]> GetExcelSheetNamesAsync(Stream stream);

    Task<Iec608705Table[]> ReadTablesFromExcelAsync(Stream stream);

    Task<Stream> GetExcelFromTablesAsync(Iec608705Table[] tables);
}