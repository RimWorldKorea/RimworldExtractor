using RimExtractorCore.DataTypes;

namespace RimExtractorCore.Extractor;

public interface IExtractionProcedure
{
    string Name { get; }
    IEnumerable<TranslationEntry> Extract(SimulationResult simResult);
}