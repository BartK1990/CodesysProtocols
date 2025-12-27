namespace CodesysProtocols.Model.TableData.Iec608705;

public class Iec608705Column
{
    public string NodeName { get; set; }

    public string Name { get; set; }

    public string Type { get; set; }

    public List<Iec608705TableValue> Values { get; set; } = new List<Iec608705TableValue>();
}
