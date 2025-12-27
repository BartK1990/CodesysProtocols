using System.Xml.Linq;
using System.Xml.Serialization;

namespace CodesysProtocols.Model.Xml.Iec608705.Nodes;

[Serializable]
[System.ComponentModel.DesignerCategory("code")]
[XmlType(AnonymousType = true)]
public class Iec608705XmlText : Iec608705XmlBase
{
    public new const string NodeName = "Text";

    public Iec608705XmlText()
    {
        base.NodeName = NodeName;
    }

    [XmlAttribute]
    public string Value { get; set; }

    public static Iec608705XmlText Create(XElement element)
    {
        return new Iec608705XmlText
        {
            Value = element.Value,
        };
    }

    protected override void SetAttributesAndValue(XElement element)
    {
        element.Value = Value;
    }
}