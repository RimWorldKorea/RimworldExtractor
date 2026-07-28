using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.DataTypes;

public class SimulationResult
{
    public ModMetadata TargetMod { get; set; } = null!;
    public List<ExtractableFolder> TargetFolders { get; set; } = new();
    public List<ModMetadata> ReferenceMods { get; set; } = new();
    
    // 기존 단일 트리 대신 스냅샷 리스트를 가집니다. (인덱스 0은 항상 Base)
    public List<DefSnapshot> Snapshots { get; set; } = new();
}