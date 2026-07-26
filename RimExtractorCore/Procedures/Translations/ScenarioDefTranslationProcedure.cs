using System.Collections.Generic;
using System.Linq;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore.Procedures;

public class ScenarioDefTranslationProcedure : ITranslationEntryProcedure
{
    public string Name => "ScenarioDefTranslationProcedure";

    public IEnumerable<TranslationEntry> Process(IEnumerable<TranslationEntry> entries)
    {
        var lst = entries.ToList();
        foreach (var entry in lst)
        {
            if (entry.ClassName.EndsWith("ScenarioDef") && entry.RealNode is "scenario.name")
            {
                if (!lst.Any(x => x.ClassName.EndsWith("ScenarioDef") && x.RealNode is "label"))
                {
                    yield return entry with { Node = entry.DefName + ".label" };
                }
            }
            else if (entry.ClassName.EndsWith("ScenarioDef") && entry.RealNode is "scenario.description")
            {
                if (!lst.Any(x => x.ClassName.EndsWith("ScenarioDef") && x.RealNode is "description"))
                {
                    yield return entry with { Node = entry.DefName + ".description" };
                }
            }
            else
            {
                yield return entry;
            }
        }
    }
}