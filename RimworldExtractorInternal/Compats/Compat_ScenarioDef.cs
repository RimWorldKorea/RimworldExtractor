using System.Xml.Linq;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorInternal.Compats
{
    [CompatPriority(200)]
    internal class Compat_ScenarioDef : BaseCompat
    {
        private const string selector =
            "Defs/ScenarioDef";
        public override void DoPreProcessing(XDocument doc)
        {
            var nodes = doc.SelectNodesSafe(selector);
            if (nodes == null)
                return;

            foreach (var node in nodes)
            {
                var label = node.Element("label");
                var description = node.Element("description");

                if (label == null || description == null)
                {
                    Log.Wrn($"ScenarioDef에 label이나 description 태그가 없습니다. defName: {node.Element("defName")?.Value}");
                    continue;
                }

                if (node.Element("scenario")?.Element("name") != null || node.Element("scenario")?.Element("description") != null)
                {
                    Log.Msg($"ScenarioDef에 이미 scenario.name이나 scenario.description 태그가 존재합니다. defName: {node.Element("defName")?.Value}");
                    continue;
                }

                node.AppendElement("scenario", scenario =>
                {
                    scenario.AppendElement("name", label.Value);
                    scenario.AppendElement("description", description.Value);
                });
            }
        }

        public override IEnumerable<TranslationEntry> DoPostProcessing(IEnumerable<TranslationEntry> entries)
        {
            var lst = entries.ToList();
            foreach (var entry in lst)
            {
                if (entry.ClassName.EndsWith("ScenarioDef") && entry.RealNode is "scenario.name")
                {
                    if (!lst.Any(x =>
                            x.ClassName.EndsWith("ScenarioDef") &&
                            x.RealNode is "label"))
                    {
                        yield return entry with { Node = entry.DefName + ".label" };
                    }
                }
                else if (entry.ClassName.EndsWith("ScenarioDef") && entry.RealNode is "scenario.description")
                {
                    if (!lst.Any(x =>
                            x.ClassName.EndsWith("ScenarioDef") &&
                            x.RealNode is "description"))
                    {
                        yield return entry with { Node = entry.DefName + ".description" };
                    }
                }
                yield return entry;
            }
        }
    }
}