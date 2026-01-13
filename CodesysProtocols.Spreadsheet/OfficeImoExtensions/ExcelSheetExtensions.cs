using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using System.Reflection;

namespace CodesysProtocols.Spreadsheet.OfficeImoExtensions;

/// <summary>
/// Extension methods for ExcelSheet to support additional OpenXML features
/// </summary>
public static class ExcelSheetExtensions
{
    public static ExcelSpreadsheet? GetExcelSpreadsheet(this ExcelSheet sheet)
    {
        SpreadsheetDocument? spreadsheetDocument = GetSpreadsheetDocument(sheet);
        if (spreadsheetDocument is null) return null;
        WorksheetPart? worksheetPart = GetWorksheetPart(sheet, spreadsheetDocument);
        if (worksheetPart?.Worksheet is null) return null;

        return new ExcelSpreadsheet(spreadsheetDocument, worksheetPart);
    }

    /// <summary>
    /// Gets the underlying SpreadsheetDocument from an ExcelSheet using reflection.
    /// Note: This uses reflection to access private fields since OfficeIMO doesn't expose
    /// the underlying OpenXML objects through its public API. This approach is necessary
    /// to extend OfficeIMO's functionality but creates a dependency on internal implementation details.
    /// </summary>
    private static SpreadsheetDocument? GetSpreadsheetDocument(ExcelSheet sheet)
    {
        try
        {
            // Try to get the SpreadsheetDocument directly from ExcelSheet
            var spreadsheetDocField = sheet.GetType().GetField("_spreadSheetDocument", BindingFlags.NonPublic | BindingFlags.Instance);
            if (spreadsheetDocField != null)
            {
                return spreadsheetDocField.GetValue(sheet) as SpreadsheetDocument;
            }
            
            // Fallback: Try to get via ExcelDocument
            var excelDocField = sheet.GetType().GetField("_excelDocument", BindingFlags.NonPublic | BindingFlags.Instance);
            if (excelDocField != null)
            {
                var excelDoc = excelDocField.GetValue(sheet);
                if (excelDoc != null)
                {
                    var spreadsheetDocFromExcelDoc = excelDoc.GetType().GetField("_spreadSheetDocument", BindingFlags.NonPublic | BindingFlags.Instance);
                    if (spreadsheetDocFromExcelDoc != null)
                    {
                        return spreadsheetDocFromExcelDoc.GetValue(excelDoc) as SpreadsheetDocument;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // If reflection fails (e.g., due to OfficeIMO internal changes), return null
            // The caller will handle the null case appropriately
            System.Diagnostics.Debug.WriteLine($"Failed to get SpreadsheetDocument via reflection: {ex.Message}");
        }
        
        return null;
    }

    private static WorksheetPart? GetWorksheetPart(ExcelSheet sheet, SpreadsheetDocument spreadsheetDocument)
    {
        try
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;
            if (workbookPart == null) return null;

            // Find the worksheet part by sheet name
            var sheets = workbookPart.Workbook.Sheets;
            var sheetElement = sheets?.Elements<Sheet>().FirstOrDefault(s => s.Name == sheet.Name);
            
            if (sheetElement?.Id?.Value != null)
            {
                return workbookPart.GetPartById(sheetElement.Id.Value) as WorksheetPart;
            }
        }
        catch (Exception ex)
        {
            // Log and return null if we can't find the worksheet part
            System.Diagnostics.Debug.WriteLine($"Failed to get WorksheetPart: {ex.Message}");
        }
        
        return null;
    }
}
