using System.Xml.Linq;
using System.Xml.Serialization;

namespace CodesysProtocols.Model.Xml.Iec608705.Nodes;

[Serializable]
[System.ComponentModel.DesignerCategory("code")]
[XmlType(AnonymousType = true)]
public class Iec608705ProjectDescription : Iec608705XmlBase
{
    public new const string NodeName = "IEC60870ProjectDescription";

    public Iec608705ProjectDescription()
    {
        base.NodeName = NodeName;
    }

    [XmlAttribute]
    public string Version { get; set; }

    [XmlAttribute]
    public string Guid { get; set; }

    [XmlAttribute]
    public string Generated { get; set; }

    public static Iec608705ProjectDescription Create(XElement element)
    {
        return new Iec608705ProjectDescription
        {
            Version = element.Attribute(nameof(Version)).Value,
            Guid = element.Attribute(nameof(Guid)).Value,
            Generated = element.Attribute(nameof(Generated)).Value,
        };
    }

    protected override void SetAttributesAndValue(XElement element)
    {
        element.SetAttributeValue(nameof(Version), Version);
        element.SetAttributeValue(nameof(Guid), Guid);
        element.SetAttributeValue(nameof(Generated), Generated);
    }
}