namespace CodesysProtocols.Model;

public class ExcelSheetName
{
    public ExcelSheetName()
    {
        SheetCollection = new List<string>();
    }

    public IList<string> SheetCollection { get; set; }
}
