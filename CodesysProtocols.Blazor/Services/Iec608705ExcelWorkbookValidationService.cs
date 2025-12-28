using CodesysProtocols.Model;

namespace CodesysProtocols.Blazor.Services
{
    public class Iec608705ExcelWorkbookValidationService : IIec608705ExcelWorkbookValidationService
    {
        private readonly Iec608705ExcelWorkbookValidation _iec608705ExcelWorkbookValidation;

        public Iec608705ExcelWorkbookValidationService(Iec608705ExcelWorkbookValidation iec608705ExcelWorkbookValidation)
        {
            this._iec608705ExcelWorkbookValidation = iec608705ExcelWorkbookValidation;
        }

        public async Task<bool> CheckIfContainsValidSheetsAsync(IList<string> sheetCollection)
        {
            return await _iec608705ExcelWorkbookValidation.CheckIfContainsValidSheetsAsync(sheetCollection);
        }
    }
}
