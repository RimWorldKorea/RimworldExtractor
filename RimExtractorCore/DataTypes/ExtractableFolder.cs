// DataTypes/ExtractableFolder.cs 전체 교체

using System.IO;

namespace RimExtractorCore.DataTypes;

public record ExtractableFolder
    (ModMetadata Root, string FolderName, string? RequiredPackageId, string VersionInfo = "default")
{
    public string FullPath => Path.Combine(Root.RootDir, FolderName);

    // [NEW] 상류(ModLister)에서 주입해주는 진짜 로드 폴더의 뿌리 경로
    public string LoadFolderRoot { get; init; } = string.Empty;
    
    // [NEW] 만약 주입되지 않았다면 모드 최상위 경로를 기본값으로 사용
    public string ActualLoadFolderRoot => string.IsNullOrEmpty(LoadFolderRoot) ? Root.RootDir : LoadFolderRoot;

    public override string ToString()
    {
        return $"{VersionInfo}:::{Path.GetFileName(FolderName)}" + (RequiredPackageId != null ? $"\n[조건={RequiredPackageId}]" : "");
    }
}

public class ExtractableFolderComparer : IEqualityComparer<ExtractableFolder>
{
    public bool Equals(ExtractableFolder? x, ExtractableFolder? y)
    {
        if (ReferenceEquals(x, y)) return true;
        if (ReferenceEquals(x, null)) return false;
        if (ReferenceEquals(y, null)) return false;
        if (x.GetType() != y.GetType()) return false;
        return x.FolderName == y.FolderName;
    }

    public int GetHashCode(ExtractableFolder obj)
    {
        return HashCode.Combine(obj.FolderName, obj.VersionInfo);
    }
}