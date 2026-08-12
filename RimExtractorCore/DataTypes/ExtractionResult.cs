namespace RimExtractorCore.DataTypes;

/// <summary>
/// 메타데이터와 함께 기록된 TranslationEntry 컬렉션입니다.
/// </summary>
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