using System.Xml.Linq;

namespace CodesysProtocols.Model.Xml.Iec608705.Nodes;

public abstract class Iec608705XmlBase
{
    public string NodeName { get; protected set; }

    public XElement GetElement()
    {
        var element = new XElement(NodeName);

        SetAttributesAndValue(element);

        return element;
    }

    protected abstract void SetAttributesAndValue(XElement element);
}