using System.Xml.Linq;
using System.Xml.Serialization;

namespace CodesysProtocols.Model.Xml.Iec608705.Nodes;

[Serializable]
[System.ComponentModel.DesignerCategory("code")]
[XmlType(AnonymousType = true)]
public class Iec608705XmlChannel : Iec608705XmlBase
{
    public new const string NodeName = "Channel";

    public Iec608705XmlChannel()
    {
        base.NodeName = NodeName;
    }

    [XmlAttribute]
    public string Name { get; set; }

    [XmlAttribute]
    public string Type { get; set; }

    [XmlAttribute]
    public string Autoapply { get; set; }

    [XmlAttribute]
    public string Assignment { get; set; }

    public static Iec608705XmlChannel Create(XElement element)
    {
        return new Iec608705XmlChannel
        {
            Name = element.Attribute(nameof(Name)).Value,
            Type = element.Attribute(nameof(Type)).Value,
            Autoapply = element.Attribute(nameof(Autoapply)).Value,
            Assignment = element.Attribute(nameof(Assignment)).Value,
        };
    }

    protected override void SetAttributesAndValue(XElement element)
    {
        element.SetAttributeValue(nameof(Name), Name);
        element.SetAttributeValue(nameof(Type), Type);
        element.SetAttributeValue(nameof(Autoapply), Autoapply);
        element.SetAttributeValue(nameof(Assignment), Assignment);
    }
}