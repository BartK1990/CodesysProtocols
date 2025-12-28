using System.Xml.Linq;
using System.Xml.Serialization;

namespace CodesysProtocols.Model.Xml.Iec608705.Nodes;

public class Iec608705Node : Iec608705XmlBase
{
    [XmlAttribute]
    public string Name { get; set; }

    public void SetNodeName(string name)
    {
        NodeName = name;
    }

    public static Iec608705Node Create(XElement element)
    {
        return new Iec608705Node
        {
            Name = element.Attribute(nameof(Name)).Value,
        };
    }

    protected override void SetAttributesAndValue(XElement element)
    {
        element.SetAttributeValue(nameof(Name), Name);
    }
}