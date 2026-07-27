using RimExtractorCore.DataTypes;

namespace RimExtractorCore.Procedures;

public interface IExtractionProcedure
{
    string Name { get; }
    IEnumerable<TranslationEntry> Extract(SimulationResult simResult);
}