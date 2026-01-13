using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using System.Reflection;

namespace CodesysProtocols.Spreadsheet.OfficeImoExtensions;

public static class ExcelDocumentExtensions
{
    extension(ExcelDocument excelDocument)
    {
        public static ExcelDocument Create(Stream stream)
        {
            var obj = new ExcelDocument();

            var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = (obj._spreadSheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)).AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            SetPrivateField(obj, "_workBookPart", workbookPart);

            obj.BuiltinDocumentProperties = new BuiltinDocumentProperties(obj);
            obj.ApplicationProperties = new ApplicationProperties(obj);

            return obj;
        }
    }

    private static void SetPrivateField(object target, string fieldName, object value)
    {
        var field = target.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
        if (field is null)
        {
            throw new MissingFieldException(target.GetType().FullName, fieldName);
        }

        field.SetValue(target, value);
    }
}
