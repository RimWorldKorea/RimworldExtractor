using System.Xml.Linq;
using RimworldExtractorInternal.Core;

namespace RimworldExtractorInternal.DefTree.Procedures;

public class ScenarioDefProcedure : IXDocumentProcedure
{
    public string Name => "ScenarioDefProcedure";

    public PipelineStage Stage => PipelineStage.StageB;

    public XDocument Process(XDocument defTree)
    {
        var selector = "Defs/ScenarioDef";
        var nodes = defTree.SelectNodesSafe(selector);
        if (nodes == null) return defTree;

        foreach (var node in nodes)
        {
            var label = node.Element("label");
            var description = node.Element("description");

            if (label == null || description == null)
            {
                Log.Wrn($"ScenarioDef에 label 또는 description이 없습니다. defName: {node.Element("defName")?.Value}");
                continue;
            }

            if (node.Element("scenario")?.Element("name") != null || node.Element("scenario")?.Element("description") != null)
            {
                Log.Msg($"ScenarioDef에 이미 scenario.name 또는 scenario.description이 존재합니다. defName: {node.Element("defName")?.Value}");
                continue;
            }

            node.AppendElement("scenario", scenario =>
            {
                scenario.AppendElement("name", label.Value);
                scenario.AppendElement("description", description.Value);
            });
        }

        return defTree;
    }
}