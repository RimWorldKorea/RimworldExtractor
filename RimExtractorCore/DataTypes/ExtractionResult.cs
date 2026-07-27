namespace RimExtractorCore.DataTypes;

/// <summary>
/// TranslationEntry 컬렉션을 메타데이터와 함께 기록합니다.
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