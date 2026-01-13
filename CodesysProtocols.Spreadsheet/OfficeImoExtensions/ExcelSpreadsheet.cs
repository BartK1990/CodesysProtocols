using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;

namespace CodesysProtocols.Spreadsheet.OfficeImoExtensions;

public record ExcelSpreadsheet(SpreadsheetDocument SpreadsheetDocument, WorksheetPart WorksheetPart)
{
    public void CellBorder(
        int row,
        int col,
        BorderStyle topStyle = BorderStyle.None,
        BorderStyle bottomStyle = BorderStyle.None,
        BorderStyle leftStyle = BorderStyle.None,
        BorderStyle rightStyle = BorderStyle.None,
        string? color = null,
        bool save = true)
    {
        if (WorksheetPart.Worksheet is null) return;
        var worksheet = WorksheetPart.Worksheet;
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;

        // Get or create the cell
        var cell = GetOrCreateCell(worksheet, sheetData, row, col);

        // Get or create stylesheet
        var styleSheet = GetOrCreateStylesheet(SpreadsheetDocument);

        // Create or get border with specified styles
        var borderId = CreateBorder(styleSheet, topStyle, bottomStyle, leftStyle, rightStyle, color);

        // Apply the border to the cell
        ApplyCellFormat(cell, styleSheet, borderId);

        if (save)
        {
            WorksheetPart.Worksheet.Save();
            SpreadsheetDocument.WorkbookPart?.WorkbookStylesPart?.Stylesheet?.Save();
        }
    }

    public void Save()
    {
        WorksheetPart.Worksheet?.Save();
        SpreadsheetDocument.WorkbookPart?.WorkbookStylesPart?.Stylesheet?.Save();
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
        var workbookPart = spreadsheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
        var workbookStylesPart = workbookPart.WorkbookStylesPart;
        if (workbookStylesPart?.Stylesheet is null)
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
                var styleIndex = (int)cell.StyleIndex.Value;
                var childElements = cellFormats.ChildElements;
                if (styleIndex >= 0 && styleIndex < childElements.Count)
                {
                    if (childElements[styleIndex] is CellFormat existingCellFormat)
                    {
                        cellFormat.FontId = existingCellFormat.FontId;
                        cellFormat.FillId = existingCellFormat.FillId;
                        cellFormat.NumberFormatId = existingCellFormat.NumberFormatId;
                        cellFormat.ApplyFont = existingCellFormat.ApplyFont;
                        cellFormat.ApplyFill = existingCellFormat.ApplyFill;
                        cellFormat.ApplyNumberFormat = existingCellFormat.ApplyNumberFormat;
                    }
                }
            }

            cellFormats.Append(cellFormat);
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;
            formatId = cellFormats.Count.Value - 1;
        }

        cell.StyleIndex = formatId;
    }
}

