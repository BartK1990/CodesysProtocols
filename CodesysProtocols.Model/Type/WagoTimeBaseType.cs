namespace CodesysProtocols.Model.Type;

public enum WagoTimeBaseTypes
{
    Millisecond,
    Second,
    Minute,
    Hour,
    Day
}

public static class WagoTimeBaseType
{
    public static readonly Dictionary<string, string> TypesDict = new Dictionary<string, string>
        {
            { WagoTimeBaseTypes.Millisecond.ToString(), "ms" },
            { WagoTimeBaseTypes.Second.ToString(), "s" },
            { WagoTimeBaseTypes.Minute.ToString(), "m" },
            { WagoTimeBaseTypes.Hour.ToString(), "h" },
            { WagoTimeBaseTypes.Day.ToString(), "d" }
        };
}
