using System.Xml.Linq;
using System.Xml.Serialization;

namespace CodesysProtocols.Model.Xml.Iec608705.Nodes;

[Serializable]
[System.ComponentModel.DesignerCategory("code")]
[XmlType(AnonymousType = true)]
public class Iec608705XmlParameter : Iec608705XmlBase
{
    public new const string NodeName = "Parameter";

    public Iec608705XmlParameter()
    {
        base.NodeName = NodeName;
    }

    [XmlAttribute]
    public string Name { get; set; }

    [XmlAttribute]
    public string Type { get; set; }

    [XmlAttribute]
    public string Val { get; set; }

    [XmlAttribute]
    public string Frame { get; set; }

    public static Iec608705XmlParameter Create(XElement element)
    {
        return new Iec608705XmlParameter
        {
            Name = element.Attribute(nameof(Name)).Value,
            Type = element.Attribute(nameof(Type)).Value,
            Val = element.Attribute(nameof(Val)).Value,
            Frame = element.Attribute(nameof(Frame))?.Value, // This one is optional
        };
    }

    protected override void SetAttributesAndValue(XElement element)
    {
        element.SetAttributeValue(nameof(Name), Name);
        element.SetAttributeValue(nameof(Type), Type);
        element.SetAttributeValue(nameof(Val), Val);
        element.SetAttributeValue(nameof(Frame), Frame);
    }
}