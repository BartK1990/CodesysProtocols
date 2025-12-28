namespace CodesysProtocols.DataAccess.Entities;

public class Codesys23Iec608705Counter
{
    public int Id { get; set; }

    public required string Name { get; set; }

    public long Value { get; set; }
}