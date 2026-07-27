using System.Xml.Linq;
using RimExtractorCore;
using RimExtractorCore.Extractor;

namespace RimExtractorCore.Procedures;

public class VerbProcedure : IXDocumentProcedure
{
    public string Name => "VerbProcedure";

    public InjectionStage Stage => InjectionStage.StageB;

    public XDocument Process(XDocument defTree)
    {
        var selector = "Defs/ThingDef/verbs/*[.//verbClass[contains(text(), 'Verb_Shoot') or contains(text(), 'Verb_ShootOneUse') or contains(text(), 'Verb_ShootWithSmoke')]]";
        var nodes = defTree.SelectNodesSafe(selector);
        if (nodes == null) return defTree;

        foreach (var node in nodes)
        {
            var root = ExtractorEngine.GetRootDefNode(node, out _);
            if (root == null || root.HasAttribute("Abstract"))
                continue;

            var labelNode = root.Element("label");
            if (labelNode == null)
            {
                Log.Wrn($"Abstract가 아닌 Def에 label이 존재하지 않습니다. {node}");
                continue;
            }

            var verbLabelNode = node.Element("label");
            if (verbLabelNode == null)
            {
                node.AppendElement("label", labelNode.Value);
            }
        }

        return defTree;
    }
}