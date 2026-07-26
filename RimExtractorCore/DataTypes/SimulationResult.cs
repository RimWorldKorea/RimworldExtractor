using System.Xml.Linq;

namespace RimExtractorCore.DataTypes;

public class SimulationResult
{
    public XDocument DefTree { get; set; } = new(new XElement("Defs"));
    public Dictionary<string, XElement> ParentNodeLookUp { get; } = new();
    public List<(RequiredMods? RequiredMods, XElement Element)> DefsAddedByPatches { get; } = new();
        
    public ModMetadata? TargetMod { get; set; }
    public List<ExtractableFolder> TargetFolders { get; set; } = new();
    public List<ModMetadata> ReferenceMods { get; set; } = new();
}