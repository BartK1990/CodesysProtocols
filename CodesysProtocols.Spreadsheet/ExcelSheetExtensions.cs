using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using System.Reflection;

namespace CodesysProtocols.Spreadsheet;

/// <summary>
/// Extension methods for ExcelSheet to support additional OpenXML features
/// </summary>
public static class ExcelSheetExtensions
{
    /// <summary>
    /// Border style options for Excel cells
    /// </summary>
    public enum BorderStyle
    {
        None,
        Thin,
        Medium,
        Thick,
        Double,
        Dotted,
        Dashed
    }

    /// <summary>
    /// Applies border to a single cell using OpenXML
    /// </summary>
    /// <param name="sheet">The Excel sheet</param>
    /// <param name="row">Row number (1-based)</param>
    /// <param name="col">Column number (1-based)</param>
    /// <param name="topStyle">Top border style</param>
    /// <param name="bottomStyle">Bottom border style</param>
    /// <param name="leftStyle">Left border style</param>
    /// <param name="rightStyle">Right border style</param>
    /// <param name="color">Border color in hex format (e.g., "000000" for black)</param>
    public static void CellBorder(
        this ExcelSheet sheet,
        int row,
        int col,
        BorderStyle topStyle = BorderStyle.None,
        BorderStyle bottomStyle = BorderStyle.None,
        BorderStyle leftStyle = BorderStyle.None,
        BorderStyle rightStyle = BorderStyle.None,
        string? color = null)
    {
        // Get the underlying OpenXML objects using reflection
        var spreadsheetDocument = GetSpreadsheetDocument(sheet);
        if (spreadsheetDocument == null) return;

        var worksheetPart = GetWorksheetPart(sheet, spreadsheetDocument);
        if (worksheetPart == null) return;

        var worksheet = worksheetPart.Worksheet;
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;

        // Get or create the cell
        var cell = GetOrCreateCell(worksheet, sheetData, row, col);
        
        // Get or create stylesheet
        var styleSheet = GetOrCreateStylesheet(spreadsheetDocument);
        
        // Create or get border with specified styles
        var borderId = CreateBorder(styleSheet, topStyle, bottomStyle, leftStyle, rightStyle, color);
        
        // Apply the border to the cell
        ApplyCellFormat(cell, styleSheet, borderId);
        
        // Save changes
        worksheetPart.Worksheet.Save();
        spreadsheetDocument.WorkbookPart?.WorkbookStylesPart?.Stylesheet.Save();
    }

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
        catch
        {
            // Fallback - return null if reflection fails
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
        catch
        {
            // Return null if we can't find the worksheet part
        }
        
        return null;
    }

    private static Cell GetOrCreateCell(Worksheet worksheet, SheetData sheetData, int rowIndex, int columnIndex)
    {
        var cellReference = $"{A1.ColumnIndexToLetters(columnIndex)}{rowIndex}";
        
        // Find or create the row
        Row? row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = (uint)rowIndex };
            sheetData.Append(row);
        }

        // Find or create the cell
        Cell? cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellReference);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellReference };
            row.Append(cell);
        }

        return cell;
    }

    private static Stylesheet GetOrCreateStylesheet(SpreadsheetDocument spreadsheetDocument)
    {
        var workbookPart = spreadsheetDocument.WorkbookPart;
        if (workbookPart == null)
            throw new InvalidOperationException("WorkbookPart is null");

        var workbookStylesPart = workbookPart.WorkbookStylesPart;
        if (workbookStylesPart == null)
        {
            workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = new Stylesheet(
                new Fonts(new Font()),
                new Fills(new Fill(), new Fill()),
                new Borders(new Border()),
                new CellFormats(new CellFormat())
            );
        }

        return workbookStylesPart.Stylesheet;
    }

    private static uint CreateBorder(
        Stylesheet styleSheet,
        BorderStyle topStyle,
        BorderStyle bottomStyle,
        BorderStyle leftStyle,
        BorderStyle rightStyle,
        string? color)
    {
        var borders = styleSheet.Borders;
        if (borders == null)
        {
            borders = new Borders();
            styleSheet.Borders = borders;
        }

        var border = new Border();
        
        if (leftStyle != BorderStyle.None)
        {
            border.LeftBorder = CreateBorderElement<LeftBorder>(leftStyle, color);
        }
        
        if (rightStyle != BorderStyle.None)
        {
            border.RightBorder = CreateBorderElement<RightBorder>(rightStyle, color);
        }
        
        if (topStyle != BorderStyle.None)
        {
            border.TopBorder = CreateBorderElement<TopBorder>(topStyle, color);
        }
        
        if (bottomStyle != BorderStyle.None)
        {
            border.BottomBorder = CreateBorderElement<BottomBorder>(bottomStyle, color);
        }

        borders.Append(border);
        borders.Count = (uint)borders.ChildElements.Count;

        return borders.Count.Value - 1;
    }

    private static T CreateBorderElement<T>(BorderStyle style, string? color) where T : BorderPropertiesType, new()
    {
        var border = new T
        {
            Style = ConvertBorderStyle(style)
        };

        if (!string.IsNullOrEmpty(color))
        {
            border.Color = new Color { Rgb = new HexBinaryValue(color) };
        }

        return border;
    }

    private static BorderStyleValues ConvertBorderStyle(BorderStyle style)
    {
        return style switch
        {
            BorderStyle.Thin => BorderStyleValues.Thin,
            BorderStyle.Medium => BorderStyleValues.Medium,
            BorderStyle.Thick => BorderStyleValues.Thick,
            BorderStyle.Double => BorderStyleValues.Double,
            BorderStyle.Dotted => BorderStyleValues.Dotted,
            BorderStyle.Dashed => BorderStyleValues.Dashed,
            _ => BorderStyleValues.None
        };
    }

    private static void ApplyCellFormat(Cell cell, Stylesheet styleSheet, uint borderId)
    {
        var cellFormats = styleSheet.CellFormats;
        if (cellFormats == null)
        {
            cellFormats = new CellFormats();
            styleSheet.CellFormats = cellFormats;
        }

        // Check if we need to create a new format or can reuse existing
        var existingFormat = cellFormats.Elements<CellFormat>()
            .FirstOrDefault(cf => cf.BorderId?.Value == borderId);

        uint formatId;
        if (existingFormat != null)
        {
            formatId = (uint)cellFormats.ChildElements.ToList().IndexOf(existingFormat);
        }
        else
        {
            var cellFormat = new CellFormat
            {
                BorderId = borderId,
                ApplyBorder = true
            };

            // Preserve existing font and fill if the cell already has a style
            if (cell.StyleIndex?.Value != null)
            {
                var existingCellFormat = cellFormats.ElementAt((int)cell.StyleIndex.Value) as CellFormat;
                if (existingCellFormat != null)
                {
                    cellFormat.FontId = existingCellFormat.FontId;
                    cellFormat.FillId = existingCellFormat.FillId;
                    cellFormat.NumberFormatId = existingCellFormat.NumberFormatId;
                    cellFormat.ApplyFont = existingCellFormat.ApplyFont;
                    cellFormat.ApplyFill = existingCellFormat.ApplyFill;
                    cellFormat.ApplyNumberFormat = existingCellFormat.ApplyNumberFormat;
                }
            }

            cellFormats.Append(cellFormat);
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;
            formatId = cellFormats.Count.Value - 1;
        }

        cell.StyleIndex = formatId;
    }
}
