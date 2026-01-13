namespace CodesysProtocols.Spreadsheet.OfficeImoExtensions;

public record ExcelRange(
    int FirstRow,
    int FirstCol,
    int LastRow,
    int LastCol);

