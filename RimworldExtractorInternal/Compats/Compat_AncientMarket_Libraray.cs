using System.Xml.Linq;

namespace RimworldExtractorInternal.Compats
{
    /// <summary>
    /// 어반루인 모드 호환성
    /// </summary>
    internal class Compat_AncientMarket_Libraray : BaseCompat
    {
        public const string selector = "Defs/AncientMarket_Libraray.CustomMapDataDef";

        public override void DoPreProcessing(XDocument doc)
        {
            var nodes = doc.SelectNodesSafe(selector);
            if (nodes == null) return;

            foreach (var node in nodes)
            {
                //ProcessNodeRecursive(node);
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
        }
    }
}