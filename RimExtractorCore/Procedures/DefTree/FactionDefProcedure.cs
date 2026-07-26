using System.Xml.Linq;
using RimworldExtractorInternal.Core;

namespace RimworldExtractorInternal.DefTree.Procedures;

public class FactionDefProcedure : IXDocumentProcedure
{
    public string Name => "FactionDefProcedure";
    
    // Defs 데이터 로드 완료 직후, 상속/패치 적용 전 실행
    public PipelineStage Stage => PipelineStage.StageB;

    public XDocument Process(XDocument defTree)
    {
        var selector = "Defs/FactionDef";
        var nodes = defTree.SelectNodesSafe(selector);
        if (nodes == null) return defTree;

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

        return defTree;
    }
}