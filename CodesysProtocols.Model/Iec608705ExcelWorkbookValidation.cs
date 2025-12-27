using CodesysProtocols.Model.Xml.Iec608705.Nodes;
using System.Collections.Immutable;

namespace CodesysProtocols.Model;

public class Iec608705ExcelWorkbookValidation
{
    public static readonly ImmutableHashSet<string> SheetNames = ImmutableHashSet.Create(
        Iec608705ProjectDescription.NodeName,
        Iec608705Converter.ConfigurationNodeName,
        Iec608705Converter.ClientsServersSheetName,
        Iec608705Converter.ConnectionsSheetName,
        Iec608705Converter.ObjectsSheetName);

    public bool CheckIfContainsValidSheets(IList<string> sheetCollection)
    {
        if (sheetCollection is null)
        {
            return false;
        }

        if (!SheetNames.Except(sheetCollection).Any())
        {
            return true;
        }

        return false;
    }
    public async Task<bool> CheckIfContainsValidSheetsAsync(IList<string> sheetCollection)
    {
        return await Task.Run(() => CheckIfContainsValidSheets(sheetCollection));
    }
}