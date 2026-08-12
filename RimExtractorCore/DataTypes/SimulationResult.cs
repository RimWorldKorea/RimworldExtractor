using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.DataTypes;

/// <summary>
/// DefTreeSimulator가 생성하는 XML 트리와 메타데이터입니다.
/// </summary>
public class SimulationResult
{
    /// <summary>
    /// 
    /// </summary>
    public ModMetadata TargetMod { get; set; } = null!;
    
    public List<DefSnapshot> Snapshots { get; set; } = new();
}