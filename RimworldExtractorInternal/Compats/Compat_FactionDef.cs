using System.Xml.Linq;

namespace RimworldExtractorInternal.Compats
{
    internal class Compat_FactionDef : BaseCompat
    {
        private const string selector =
            "Defs/FactionDef";
        public override void DoPreProcessing(XDocument doc)
        {
            var nodes = doc.SelectNodesSafe(selector);
            if (nodes == null)
                return;
            foreach (var node in nodes)
            {
                var pawnSingular = node.Element("pawnSingular");
                var pawnsPlural = node.Element("pawnsPlural");
                var leaderTitle = node.Element("leaderTitle");

                if (pawnSingular == null)
                    node.AppendElement("pawnSingular", "member");
                if (pawnsPlural == null)
                    node.AppendElement("pawnsPlural", "members");
                if (leaderTitle == null)
                    node.AppendElement("leaderTitle", "leader");
            }
        }
    }
}
