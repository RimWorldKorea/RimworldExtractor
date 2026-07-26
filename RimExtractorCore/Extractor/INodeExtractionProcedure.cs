using System.Collections.Generic;
using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.Extractor;

public interface INodeExtractionProcedure
{
    string Name { get; }
    IEnumerable<TranslationEntry> Extract(SimulationResult simResult);
}