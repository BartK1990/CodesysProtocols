using CodesysProtocols.DataAccess.Enums;
using CodesysProtocols.DataAccess.Services;
using CodesysProtocols.Model;
using CodesysProtocols.Model.TableData.Iec608705;
using System.Xml.Linq;

namespace CodesysProtocols.Blazor.Services.Codesys23;

public class Iec608705ConverterService : IIec608705ConverterService
{
    private readonly Iec608705Converter _iec608705Converter;

    private readonly IIec608705CounterService _iec608705CounterService;

    public Iec608705ConverterService(Iec608705Converter iec608705Converter, IIec608705CounterService iec608705CounterService)
    {
        _iec608705Converter = iec608705Converter;
        _iec608705CounterService = iec608705CounterService;
    }

    public async Task<XDocument> DataToXmlAsync(Iec608705Table[] tables)
    {
        await _iec608705CounterService.IncrementCounterValueAsync(Codesys23Iec608705CounterType.ToXml);
        return _iec608705Converter.DataToXml(tables);
    }

    public async Task<Iec608705Table[]> XmlToDataAsync(XDocument xml)
    {
        await _iec608705CounterService.IncrementCounterValueAsync(Codesys23Iec608705CounterType.ToExcel);
        return _iec608705Converter.XmlToData(xml);
    }
}