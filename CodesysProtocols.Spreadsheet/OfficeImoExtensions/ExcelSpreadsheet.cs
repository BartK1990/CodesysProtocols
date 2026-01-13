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

    /// <summary>
    /// Auto-fits columns based on content length without relying on font information.
    /// This implementation works in containerized environments where font information may not be available.
    /// </summary>
    /// <param name="minWidth">Minimum column width (default: 8)</param>
    /// <param name="maxWidth">Maximum column width (default: 50)</param>
    /// <param name="paddingChars">Extra characters to add for padding (default: 2)</param>
    public void AutoFitColumns(double minWidth = 8, double maxWidth = 50, int paddingChars = 2)
    {
        if (WorksheetPart.Worksheet is null) return;
        
        var worksheet = WorksheetPart.Worksheet;
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;

        // Get or create Columns element
        var columns = worksheet.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            worksheet.InsertBefore(columns, sheetData);
        }

        // Calculate max content length for each column
        var columnWidths = new Dictionary<int, double>();
        
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value == null) continue;
                
                var columnIndex = GetColumnIndex(cell.CellReference.Value);
                var cellText = GetCellText(cell);
                var contentLength = cellText?.Length ?? 0;
                
                if (!columnWidths.ContainsKey(columnIndex) || columnWidths[columnIndex] < contentLength)
                {
                    columnWidths[columnIndex] = contentLength;
                }
            }
        }

        // Apply calculated widths to columns
        foreach (var kvp in columnWidths)
        {
            var columnIndex = kvp.Key;
            var contentLength = kvp.Value;
            
            // Calculate width: content length + padding, clamped to min/max
            // Excel width units are roughly character widths, with some adjustments
            var width = Math.Max(minWidth, Math.Min(maxWidth, contentLength + paddingChars));
            
            // Check if column already exists
            var existingColumn = columns.Elements<Column>()
                .FirstOrDefault(c => c.Min?.Value == (uint)columnIndex && c.Max?.Value == (uint)columnIndex);
            
            if (existingColumn != null)
            {
                existingColumn.Width = width;
                existingColumn.CustomWidth = true;
            }
            else
            {
                var column = new Column
                {
                    Min = (uint)columnIndex,
                    Max = (uint)columnIndex,
                    Width = width,
                    CustomWidth = true
                };
                columns.Append(column);
            }
        }

        Save();
    }

    private static int GetColumnIndex(string cellReference)
    {
        // Extract column letters from cell reference (e.g., "A1" -> "A", "AB10" -> "AB")
        var columnLetters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        
        // Convert column letters to index (A=1, B=2, ..., Z=26, AA=27, etc.)
        int columnIndex = 0;
        for (int i = 0; i < columnLetters.Length; i++)
        {
            columnIndex = columnIndex * 26 + (columnLetters[i] - 'A' + 1);
        }
        
        return columnIndex;
    }

    private string? GetCellText(Cell cell)
    {
        if (cell.CellValue == null) return null;
        
        var value = cell.CellValue.Text;
        
        // If it's a shared string, look it up in the shared string table
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var stringTable = SpreadsheetDocument.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
            if (stringTable != null && int.TryParse(value, out int index))
            {
                var sharedString = stringTable.Elements<SharedStringItem>().ElementAtOrDefault(index);
                return sharedString?.InnerText;
            }
        }
        
        return value;
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

