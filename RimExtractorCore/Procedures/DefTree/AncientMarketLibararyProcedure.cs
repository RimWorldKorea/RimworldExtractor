using System.Collections.Generic;
using System.Xml.Linq;
using RimExtractorCore;

namespace RimExtractorCore.Procedures;

public class AncientMarketLibraryProcedure : IXDocumentProcedure
{
    public string Name => "AncientMarketLibraryProcedure";

    public InjectionStage Stage => InjectionStage.StageB;

    public XDocument Process(XDocument defTree)
    {
        var selector = "Defs/AncientMarket_Libraray.CustomMapDataDef";
        var nodes = defTree.SelectNodesSafe(selector);
        if (nodes == null) return defTree;

        foreach (var node in nodes)
        {
            var q = new Queue<XElement>();
            q.Enqueue(node);
            while (q.Count > 0)
            {
                var n = q.Dequeue();
                var toRemove = new List<XElement>();
                foreach (var child in n.Elements())
                {
                    if (child.IsTextNode() && child.Value.StartsWith('(') && child.Value.EndsWith(')'))
                    {
                        toRemove.Add(child);
                    }
                    else
                    {
                        q.Enqueue(child);
                    }
                }
                foreach (var child in toRemove)
                {
                    child.Remove();
                }
            }
        }

        return defTree;
    }
}