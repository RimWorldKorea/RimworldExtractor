using System.Collections.Generic;

namespace RimExtractorCore.DataTypes;

public class ExtractionResult
{
    public ModMetadata TargetMod { get; }
    public IReadOnlyList<TranslationEntry> Entries { get; }

    public ExtractionResult(ModMetadata targetMod, List<TranslationEntry> entries)
    {
        TargetMod = targetMod;
        Entries = entries.AsReadOnly();
    }
}