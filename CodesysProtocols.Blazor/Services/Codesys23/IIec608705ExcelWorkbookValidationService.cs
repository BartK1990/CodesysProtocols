namespace CodesysProtocols.Blazor.Services.Codesys23
{
    public interface IIec608705ExcelWorkbookValidationService
    {
        Task<bool> CheckIfContainsValidSheetsAsync(IList<string> sheetCollection);
    }
}
