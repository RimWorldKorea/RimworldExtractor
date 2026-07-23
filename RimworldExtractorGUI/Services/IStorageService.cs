namespace RimworldExtractorGUI.Services;

public interface IStorageService
{
    Task<string?> OpenFileAsync(string title, string filterName, params string[] patterns);
    Task<string[]?> OpenFilesAsync(string title, string filterName, params string[] patterns);
    Task<string?> OpenFolderAsync(string title);
    Task<string[]?> OpenFoldersAsync(string title);
    Task<string?> SaveFileAsync(string title, string defaultFileName, string extension);
}