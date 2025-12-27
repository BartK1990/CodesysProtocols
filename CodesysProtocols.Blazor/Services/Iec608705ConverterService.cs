using CodesysProtocols.Model;
using CodesysProtocols.Model.TableData.Iec608705;
using System.Xml.Linq;

namespace CodesysProtocols.Blazor.Services;

public class Iec608705ConverterService : IIec608705ConverterService
{
    private readonly Iec608705Converter _iec608705Converter;

    public Iec608705ConverterService(Iec608705Converter iec608705Converter)
    {
        _iec608705Converter = iec608705Converter;
    }

    public async Task<XDocument> DataToXmlAsync(Iec608705Table[] tables) => await Task.Run(() => DataToXml(tables));

    private XDocument DataToXml(Iec608705Table[] tables)
    {
        return _iec608705Converter.DataToXml(tables);
    }

    public async Task<Iec608705Table[]> XmlToDataAsync(XDocument xml) => await Task.Run(() => XmlToData(xml));

    private Iec608705Table[] XmlToData(XDocument xml)
    {
        return _iec608705Converter.XmlToData(xml);
    }
}