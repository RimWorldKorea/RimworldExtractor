using System.Collections.Generic;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore;

public interface INodeExtractionProcedure
{
    string Name { get; }
    IEnumerable<TranslationEntry> Extract(SimulationResult simResult);
}