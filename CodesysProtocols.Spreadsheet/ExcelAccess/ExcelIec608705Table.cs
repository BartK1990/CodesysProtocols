using CodesysProtocols.Model.TableData.Iec608705;
using OfficeIMO.Excel;

namespace CodesysProtocols.Spreadsheet.ExcelAccess;

public class ExcelIec608705Table
{
    private const string NoValue = "---";
    private const string HeadersHtmlColor = @"#ADD8E6";

    public static Iec608705Table Read(ExcelSheetReader sheetReader, string sheetName)
    {
        var usedRange = sheetReader.GetUsedRangeA1();
        if (string.IsNullOrEmpty(usedRange))
        {
            return new Iec608705Table { Name = sheetName };
        }

        // Read the entire used range into an array
        var data = sheetReader.ReadRange(usedRange, null, default);
        int rows = data.GetLength(0);
        int columns = data.GetLength(1);

        var table = new Iec608705Table { Name = sheetName };

        for (int columnNo = 0; columnNo < columns; columnNo++)
        {
            var column = new Iec608705Column();
            table.Columns.Add(column);
            for (int rowNo = 0; rowNo < rows; rowNo++)
            {
                int tableRowNo = rowNo + 1;
                var cellValue = Excel.GetCellValue(data[rowNo, columnNo]);
                
                switch (tableRowNo)
                {
                    case 1:
                        column.NodeName = (cellValue is NoValue ? null : cellValue) ?? string.Empty;
                        break;
                    case 2:
                        column.Name = (cellValue is NoValue ? null : cellValue) ?? string.Empty;
                        break;
                    case 3:
                        column.Type = (cellValue is NoValue ? null : cellValue) ?? string.Empty;
                        break;
                    default:
                        if (cellValue is null)
                        {
                            break;
                        }

                        if (cellValue == NoValue)
                        {
                            cellValue = string.Empty;
                        }

                        column.Values.Add(new Iec608705TableValue 
                            { Row = tableRowNo, Value = cellValue });
                        break;
                }
            }
        }

        return table;
    }

    public static void Write(Iec608705Table table, ExcelSheet sheet)
    {
        for (int i = 0; i < table.Columns.Count; i++)
        {
            var column = table.Columns[i];

            sheet.CellValue(1, i + 1, column.NodeName ?? NoValue);
            sheet.CellValue(2, i + 1, column.Name ?? NoValue);
            sheet.CellValue(3, i + 1, column.Type ?? NoValue);

            foreach (var value in column.Values)
            {
                sheet.CellValue(value.Row, i + 1, string.IsNullOrEmpty(value.Value) ? NoValue : value.Value);
            }
        }

        PrepareSheet(sheet);
    }

    private static void PrepareSheet(ExcelSheet sheet)
    {
        var lastColumn = Excel.GetLastSheetColumn(sheet);
        if (lastColumn < 1)
        {
            return;
        }

        sheet.AddAutoFilter($"A3:{A1.ColumnIndexToLetters(lastColumn)}3");
        sheet.Freeze(topRows: 3, leftCols: 4);
        sheet.AutoFitColumns();

        var range = new Range(1, 1, 3, lastColumn);

        range.Apply((x, y) => sheet.CellBackground(x, y, HeadersHtmlColor));

        // Add border support by extending OfficeIMO
        range.Apply((x, y) => sheet.CellBorder(x, y, 
            topStyle: ExcelSheetExtensions.BorderStyle.Thin,
            bottomStyle: ExcelSheetExtensions.BorderStyle.Thin,
            leftStyle: ExcelSheetExtensions.BorderStyle.Thin,
            rightStyle: ExcelSheetExtensions.BorderStyle.Thin));
    }
}