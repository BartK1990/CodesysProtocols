using CodesysProtocols.Model.TableData.Iec608705;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace CodesysProtocols.Spreadsheet.ExcelAccess;

public class ExcelIec608705Table
{
    private const string NoValue = "---";

    private const string HeadersHtmlColor = @"#ADD8E6";

    public static Iec608705Table Read(ExcelWorksheet ws)
    {
        int rows = Excel.GetLastSheetRow(ws);
        int columns = Excel.GetLastSheetColumn(ws);

        var table = new Iec608705Table { Name = ws.Name };

        for (int columnNo = 1; columnNo <= columns; columnNo++)
        {
            var column = new Iec608705Column();
            table.Columns.Add(column);
            for (int rowNo = 1; rowNo <= rows; rowNo++)
            {
                switch (rowNo)
                {
                    case 1:
                        string nodeName = Excel.GetCellValue(ws.Cells[rowNo, columnNo].Value);
                        column.NodeName = nodeName == NoValue ? null : nodeName;
                        break;
                    case 2:
                        string name = Excel.GetCellValue(ws.Cells[rowNo, columnNo].Value);
                        column.Name = name == NoValue ? null : name;
                        break;
                    case 3:
                        string type = Excel.GetCellValue(ws.Cells[rowNo, columnNo].Value);
                        column.Type = type == NoValue ? null : type;
                        break;
                    default:
                        var cellValue = Excel.GetCellValue(ws.Cells[rowNo, columnNo].Value);
                        if (cellValue is null)
                        {
                            break;
                        }

                        if (cellValue == NoValue)
                        {
                            cellValue = string.Empty;
                        }

                        column.Values.Add(new Iec608705TableValue 
                            { Row = rowNo, Value = cellValue });
                        break;
                }
            }
        }

        return table;
    }

    public static void Write(Iec608705Table table, ExcelWorksheet ws)
    {
        for (int i = 0; i < table.Columns.Count; i++)
        {
            var column = table.Columns[i];

            ws.Cells[1, i + 1].Value = column.NodeName ?? NoValue;
            ws.Cells[2, i + 1].Value = column.Name ?? NoValue;
            ws.Cells[3, i + 1].Value = column.Type ?? NoValue;

            foreach (var value in column.Values)
            {
                ws.Cells[value.Row, i + 1].Value = string.IsNullOrEmpty(value.Value) ? NoValue : value.Value;
            }
        }

        PrepareSheet(ws);
    }

    private static void PrepareSheet(ExcelWorksheet ws)
    {
        var lastColumn = Excel.GetLastSheetColumn(ws);
        if (lastColumn < 1)
        {
            return;
        }

        ws.Cells[3, 1, 3, lastColumn].AutoFilter = true;
        ws.View.FreezePanes(4, 5);

        ExcelRange range = ws.Cells[1, 1, 3, lastColumn];
        range.AutoFitColumns();
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(HeadersHtmlColor));
        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
    }
}