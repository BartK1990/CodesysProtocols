namespace CodesysProtocols.Spreadsheet;

public record Range(
    int FirstRow,
    int FirstCol,
    int LastRow,
    int LastCol);

