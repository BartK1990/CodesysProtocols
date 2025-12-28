namespace CodesysProtocols.Blazor.Services
{
    public interface IIec608705ExcelWorkbookValidationService
    {
        Task<bool> CheckIfContainsValidSheetsAsync(IList<string> sheetCollection);
    }
}
