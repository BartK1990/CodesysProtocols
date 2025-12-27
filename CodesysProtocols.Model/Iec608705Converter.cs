using CodesysProtocols.Model.TableData.Iec608705;
using CodesysProtocols.Model.Xml.Enums;
using CodesysProtocols.Model.Xml.Iec608705.Nodes;
using System.Xml.Linq;

namespace CodesysProtocols.Model;

public class Iec608705Converter
{
    private const string HeaderPrefix = "#";
    private const string ParentHeader = $"{HeaderPrefix}Parent";
    private const string NodeNameHeader = $"{HeaderPrefix}NodeName";
    private const string NameHeader = $"{HeaderPrefix}Name";
    private const string IndexSeparator = "-";
    private const char ChannelSplit = '.';
    private const char ParameterFrameSplit = '|';
    private const string ConnectionNodeName = "CClientConnection104";

    public const string ConfigurationNodeName = "CConfiguration";
    public const string ClientsServersSheetName = "ClientsServers";
    public const string ConnectionsSheetName = "Connections";
    public const string ObjectsSheetName = "Objects";
    public const int FirstValueRow = 4;

    public XDocument DataToXml(Iec608705Table[] tables)
    {
        var xDocument = new XDocument(new XDeclaration("1.0", "utf-16", null));

        XElement projectDescription = ReadTableProjectDescription(tables.First(t=>t.Name is Iec608705ProjectDescription.NodeName));
        xDocument.Add(projectDescription);

        ReadTable(tables.First(t => t.Name is ConfigurationNodeName), projectDescription);
        ReadTable(tables.First(t => t.Name is ClientsServersSheetName), projectDescription);
        ReadTable(tables.First(t => t.Name is ConnectionsSheetName), projectDescription);
        ReadTable(tables.First(t => t.Name is ObjectsSheetName), projectDescription);

        ClearNodeNameSuffix(projectDescription);

        return xDocument;
    }

    private void ClearNodeNameSuffix(XElement projectDescription)
    {
        var nodesToClear = projectDescription.Descendants().Where(d => d.Name.ToString().Contains(IndexSeparator));
        foreach (var node in nodesToClear)
        {
            node.Name = node.Name.ToString().Split(IndexSeparator).First();
        }
    }

    public Iec608705Table[] XmlToData(XDocument xml)
    {
        XElement configuration = xml.Root.Descendants("CConfiguration").First();

        var tables = new List<Iec608705Table>
        {
            ReadProjectDescription(xml.Root),
            ReadConfiguration(configuration),
        };

        tables.AddRange(ReadClientsServers(configuration));

        return tables.ToArray();
    }

    private static XElement ReadTableProjectDescription(Iec608705Table table)
    {
        var projectDescription = new Iec608705ProjectDescription();

        projectDescription.Version = table.Columns.First(c => c.Name is nameof(projectDescription.Version)).Values.First().Value;
        projectDescription.Guid = table.Columns.First(c => c.Name is nameof(projectDescription.Guid)).Values.First().Value;
        projectDescription.Generated = table.Columns.First(c => c.Name is nameof(projectDescription.Generated)).Values.First().Value;

        return projectDescription.GetElement();
    }

    private static void ReadTable(Iec608705Table table, XElement root)
    {
        var firstRow = table.Columns.FirstOrDefault(c => c.Name is NodeNameHeader)?.Values.Select(v => v.Row).Min();
        var lastRow = table.Columns.FirstOrDefault(c => c.Name is NodeNameHeader)?.Values.Select(v => v.Row).Max();
        if (firstRow is null || lastRow is null)
        {
            return;
        }
        for (int i = (int)firstRow; i <= lastRow; i++)
        {
            var node = new Iec608705Node();
            ReadTableParameters(table, node, root, i);
        }
    }

    private static void ReadTableParameters(Iec608705Table table, Iec608705Node iecNode, XElement root, int row)
    {
        XElement parent = root.DescendantsAndSelf(table.Columns.First(c => c.Name is ParentHeader).Values.First(v => v.Row == row).Value).First();
        iecNode.SetNodeName(table.Columns.First(c => c.Name is NodeNameHeader).Values.First(v => v.Row == row).Value);
        iecNode.Name = table.Columns.First(c => c.Name is NameHeader).Values.First(v => v.Row == row).Value;
        XElement node = iecNode.GetElement();
        parent.Add(node);

        List<Iec608705Column> parameterColumns = table.Columns.Where(c => c.Name is null || !c.Name.StartsWith(HeaderPrefix)).ToList();
        List<Iec608705Column> channelColumns = table.Columns.Where(c => c.Name is not null && c.NodeName is Iec608705XmlChannel.NodeName).ToList();
        var processedChannelsNames = new List<string>();
        foreach (Iec608705Column column in parameterColumns)
        {
            Iec608705TableValue tableValue = column.Values.FirstOrDefault(v => v.Row == row);
            if (tableValue is null)
            {
                continue;
            }
            switch (column.NodeName)
            {
                case Iec608705XmlText.NodeName:
                    node.Add(new Iec608705XmlText
                    {
                        Value = tableValue.Value,
                    }.GetElement());
                    break;
                case Iec608705XmlParameter.NodeName:
                    string val = tableValue.Value;
                    string frame = null;
                    if (tableValue.Value is not null && tableValue.Value.Contains(ParameterFrameSplit))
                    {
                        var split = tableValue.Value.Split(ParameterFrameSplit);
                        val = split.First();
                        frame = split.Last();
                    }
                    node.Add(new Iec608705XmlParameter
                    {
                        Name = column.Name,
                        Type = column.Type,
                        Val = val,
                        Frame = frame
                    }.GetElement());
                    break;

                case Iec608705XmlChannel.NodeName:
                    var channelNameParts = column.Name.Split(ChannelSplit);
                    var channelName = channelNameParts.First();
                    if (processedChannelsNames.Contains(channelName))
                    {
                        break;
                    }
                    var channelSuffix = channelNameParts.Last();
                    var channelSecondSuffix = channelSuffix is nameof(Iec608705XmlChannel.Assignment)
                        ? nameof(Iec608705XmlChannel.Autoapply)
                        : nameof(Iec608705XmlChannel.Assignment);

                    Iec608705TableValue secondTableValue = channelColumns
                        .First(cc => cc.Name == $"{channelName}{ChannelSplit}{channelSecondSuffix}" 
                                     && cc.Type == column.Type)
                        .Values.FirstOrDefault(v => v.Row == row);

                    node.Add(new Iec608705XmlChannel
                    {
                        Name = channelName,
                        Type = column.Type,
                        Assignment = channelSuffix is nameof(Iec608705XmlChannel.Assignment) ? tableValue.Value : secondTableValue.Value,
                        Autoapply = channelSuffix is nameof(Iec608705XmlChannel.Assignment) ? secondTableValue.Value : tableValue.Value,
                    }.GetElement());
                    processedChannelsNames.Add(channelName);
                    break;
                default:
                    throw new InvalidDataException($"Node: [{column.NodeName}] is not supported");
            }
        }
    }

    private static Iec608705Table ReadProjectDescription(XElement xmlRoot)
    {
        var projDesc = Iec608705ProjectDescription.Create(xmlRoot);

        var table = new Iec608705Table
        {
            Name = Iec608705ProjectDescription.NodeName
        };

        table.Columns.Add(new Iec608705Column { 
            NodeName = Iec608705ProjectDescription.NodeName,
            Name = nameof(projDesc.Version),
            Values = new List<Iec608705TableValue> { new Iec608705TableValue() { Value = projDesc.Version, Row = FirstValueRow }}});

        table.Columns.Add(new Iec608705Column { 
            NodeName = Iec608705ProjectDescription.NodeName,
            Name = nameof(projDesc.Guid),
            Values = new List<Iec608705TableValue> { new Iec608705TableValue() { Value = projDesc.Guid, Row = FirstValueRow }}});

        table.Columns.Add(new Iec608705Column { 
            NodeName = Iec608705ProjectDescription.NodeName,
            Name = nameof(projDesc.Generated),
            Values = new List<Iec608705TableValue> { new Iec608705TableValue() { Value = projDesc.Generated, Row = FirstValueRow }}});

        return table;
    }

    private static Iec608705Table ReadConfiguration(XElement configuration)
    {
        var config = Iec608705Node.Create(configuration);

        var table = new Iec608705Table
        {
            Name = ConfigurationNodeName
        };

        ReadNodeParameters(configuration, table.Columns, FirstValueRow);

        return table;
    }

    private static List<Iec608705Table> ReadClientsServers(XElement configuration)
    {
        var tables = new List<Iec608705Table>();

        var tableClientsServers = new Iec608705Table
        {
            Name = ClientsServersSheetName
        };
        tables.Add(tableClientsServers);

        var tableConnections = new Iec608705Table
        {
            Name = ConnectionsSheetName
        };
        tables.Add(tableConnections);

        var tableObjects = new Iec608705Table
        {
            Name = ObjectsSheetName
        };
        tables.Add(tableObjects);

        IEnumerable<IGrouping<XName, XElement>> multipleNodes = configuration.Elements().Where(e => e.HasElements).GroupBy(g => g.Name);
        foreach (var nodeGroup in multipleNodes)
        {
            int index = 0;
            foreach (var xElement in nodeGroup)
            {
                xElement.Name += $"{IndexSeparator}{++index}";
            }
        }

        int rowClientsServers = FirstValueRow;
        int rowConnections = FirstValueRow;
        int rowObjects = FirstValueRow;
        foreach (var nodeClientServer in configuration.Elements().Where(e => e.HasElements))
        {
            ReadNodeParameters(nodeClientServer, tableClientsServers.Columns, rowClientsServers++);
            var connections = nodeClientServer.Elements().Where(e => e.HasElements && e.Name == ConnectionNodeName).ToList();
            if (!connections.Any())
            {
                ProcessObjects(nodeClientServer);
            }
            foreach (var nodeConnection in connections)
            {
                nodeConnection.Name += $"{IndexSeparator}{nodeConnection.Parent.Name.ToString().Split(IndexSeparator).Last()}";
                ReadNodeParameters(nodeConnection, tableConnections.Columns, rowConnections++);
                ProcessObjects(nodeConnection);
            }
        }

        return tables;

        void ProcessObjects(XElement nodeParent)
        {
            foreach (var nodeObject in nodeParent.Elements().Where(e => e.HasElements))
            {
                ReadNodeParameters(nodeObject, tableObjects.Columns, rowObjects++);
            }
        }
    }

    private static void ReadNodeParameters(XElement nodeElement, ICollection<Iec608705Column> columns, int row)
    {
        var firstColumn = columns.FirstOrDefault(n => n.Name == ParentHeader);
        if (firstColumn is null)
        {
            firstColumn = new Iec608705Column
            {
                Name = ParentHeader,
            };
            columns.Add(firstColumn);
        }
        firstColumn.Values.Add(new Iec608705TableValue { Value = nodeElement.Parent.Name.ToString(), Row = row });
            
        var secondColumn = columns.FirstOrDefault(n => n.Name == NodeNameHeader);
        if (secondColumn is null)
        {
            secondColumn = new Iec608705Column
            {
                Name = NodeNameHeader,
            };
            columns.Add(secondColumn);
        }
        secondColumn.Values.Add(new Iec608705TableValue { Value = nodeElement.Name.ToString(), Row = row });

        var mainNode = Iec608705Node.Create(nodeElement);
        var thirdColumn = columns.FirstOrDefault(n => n.Name == NameHeader);
        if (thirdColumn is null)
        {
            thirdColumn = new Iec608705Column
            {
                Name = NameHeader,
            };
            columns.Add(thirdColumn);
        }
        thirdColumn.Values.Add(new Iec608705TableValue { Value = mainNode.Name, Row = row });

        var elements = nodeElement.Elements().Where(e => !e.HasElements);
        foreach (var element in elements)
        {
            switch (element.Name.ToString())
            {
                case nameof(Iec608705XmlNodeNames.Parameter):
                    ProcessParameterNode(columns, row, element);
                    break;
                case nameof(Iec608705XmlNodeNames.Channel):
                    ProcessChannelNode(columns, row, element);
                    break;
                case nameof(Iec608705XmlNodeNames.Text):
                    ProcessTextNode(columns, row, element);
                    break;
                default:
                    throw new InvalidDataException();
            }
        }
    }

    private static void ProcessParameterNode(ICollection<Iec608705Column> columns, int row, XElement element)
    {
        var parameterNode = Iec608705XmlParameter.Create(element);
        var column = columns.FirstOrDefault(n => n.Name == parameterNode.Name
                                                 && n.Type == parameterNode.Type
                                                 && !n.Values.Select(v => v.Row).Contains(row));

        if (column is null)
        {
            column = new Iec608705Column
            {
                NodeName = Iec608705XmlParameter.NodeName,
                Name = parameterNode.Name,
                Type = parameterNode.Type,
            };
            columns.Add(column);
        }

        string value = parameterNode.Frame is null
            ? parameterNode.Val
            : $"{parameterNode.Val}{ParameterFrameSplit}{parameterNode.Frame}";

        column.Values.Add(new Iec608705TableValue { Value = value, Row = row });
    }

    private static void ProcessChannelNode(ICollection<Iec608705Column> columns, int row, XElement element)
    {
        var channelNode = Iec608705XmlChannel.Create(element);
        var columnsWithSplitChar = columns.Where(c => c.Name != null && c.Name.Contains(ChannelSplit)).ToList();

        ColumnSplitPart(nameof(Iec608705XmlChannel.Assignment), channelNode.Assignment);
        ColumnSplitPart(nameof(Iec608705XmlChannel.Autoapply), channelNode.Autoapply);

        void ColumnSplitPart(string suffix, string value)
        {
            var column = columnsWithSplitChar
                .FirstOrDefault(n => n.Name.Split(ChannelSplit).First() == channelNode.Name
                                     && n.Name.Split(ChannelSplit).Last() == suffix
                                     && n.Type == channelNode.Type
                                     && !n.Values.Select(v => v.Row).Contains(row));
            if (column is null)
            {
                column = new Iec608705Column
                {
                    NodeName = Iec608705XmlChannel.NodeName,
                    Name = $"{channelNode.Name}{ChannelSplit}{suffix}",
                    Type = channelNode.Type,
                };
                columns.Add(column);
            }

            column.Values.Add(new Iec608705TableValue { Value = value, Row = row });
        }
    }

    private static void ProcessTextNode(ICollection<Iec608705Column> columns, int row, XElement element)
    {
        var textNode = Iec608705XmlText.Create(element);
        var column = columns.FirstOrDefault(n => n.NodeName == Iec608705XmlText.NodeName 
                                                 && !n.Values.Select(v=>v.Row).Contains(row));
        if (column is null)
        {
            column = new Iec608705Column
            {
                NodeName = Iec608705XmlText.NodeName,
            };
            columns.Add(column);
        }

        column.Values.Add(new Iec608705TableValue { Value = textNode.Value, Row = row });
    }
}