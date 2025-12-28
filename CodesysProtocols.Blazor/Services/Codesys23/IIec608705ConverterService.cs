using CodesysProtocols.Model.TableData.Iec608705;
using System.Xml.Linq;

namespace CodesysProtocols.Blazor.Services.Codesys23;

public interface IIec608705ConverterService
{
    Task<XDocument> DataToXmlAsync(Iec608705Table[] tables);

    Task<Iec608705Table[]> XmlToDataAsync(XDocument xml);
}