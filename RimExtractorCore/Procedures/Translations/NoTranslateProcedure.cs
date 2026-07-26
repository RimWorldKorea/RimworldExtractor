using System.Collections.Generic;
using RimExtractorCore.DataTypes;
using RimExtractorCore;

namespace RimExtractorCore.Procedures;

public class NoTranslateProcedure : ITranslationEntryProcedure
{
    public string Name => "NoTranslateProcedure";
    
    private static readonly string colorChannels = "colorChannels";

    public IEnumerable<TranslationEntry> Process(IEnumerable<TranslationEntry> entries)
    {
        foreach (var translationEntry in entries)
        {
            if (translationEntry.Node.Contains(colorChannels))
                continue;

            yield return translationEntry;
        }
    }
}