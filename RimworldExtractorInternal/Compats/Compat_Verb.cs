using System.Xml.Linq;

namespace RimworldExtractorInternal.Compats
{
    internal class Compat_Verb : BaseCompat
    {
        private const string selector =
            "Defs/ThingDef/verbs/*[.//verbClass[contains(text(), 'Verb_Shoot') or contains(text(), 'Verb_ShootOneUse') or contains(text(), 'Verb_ShootWithSmoke')]]";
        public override void DoPreProcessing(XDocument doc)
        {
            var nodes = doc.SelectNodesSafe(selector);
            if (nodes == null)
                return;
            foreach (var node in nodes)
            {
                var root = Extractor.GetRootDefNode(node, out _);
                if (root == null || root.HasAttribute("Abstract"))
                    continue;
                var labelNode = root.Element("label");
                if (labelNode == null)
                {
                    Log.Wrn($"Abstract가 아닌 Def 노드에 label 노드가 없습니다. {node}");
                    continue;
                }

                var verbLabelNode = node.Element("label");
                if (verbLabelNode == null)
                {
                    node.AppendElement("label", labelNode.Value);
                }
            }
        }
    }
}