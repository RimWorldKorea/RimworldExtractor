using System.Collections.Generic;
using RimworldExtractorInternal.DataTypes;
using RimworldExtractorInternal.Core;

namespace RimworldExtractorInternal.Procedures.Translations;

public class NoTranslateProcedure : ITranslationProcedure
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