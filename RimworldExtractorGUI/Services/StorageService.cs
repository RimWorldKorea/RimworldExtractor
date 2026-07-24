using Avalonia.Controls;
using Avalonia.Platform.Storage;

namespace RimworldExtractorGUI.Services;

public interface IStorageService
{
    Task<string?> OpenFileAsync(string title, string filterName, params string[] patterns);
    Task<string[]?> OpenFilesAsync(string title, string filterName, params string[] patterns);
    Task<string?> OpenFolderAsync(string title);
    Task<string[]?> OpenFoldersAsync(string title);
    Task<string?> SaveFileAsync(string title, string defaultFileName, string extension);
}

public class StorageService : IStorageService
{
    private readonly Window _targetWindow;

    public StorageService(Window targetWindow)
    {
        _targetWindow = targetWindow;
    }

    private IStorageProvider StorageProvider => _targetWindow.StorageProvider;

    public async Task<string?> OpenFileAsync(string title, string filterName, params string[] patterns)
    {
        var files = await OpenFilesInternalAsync(title, filterName, false, patterns);
        return files?.FirstOrDefault();
    }

    public async Task<string[]?> OpenFilesAsync(string title, string filterName, params string[] patterns)
    {
        return await OpenFilesInternalAsync(title, filterName, true, patterns);
    }

    private async Task<string[]?> OpenFilesInternalAsync(string title, string filterName, bool allowMultiple, string[] patterns)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = title,
            AllowMultiple = allowMultiple,
            FileTypeFilter = new[] { new FilePickerFileType(filterName) { Patterns = patterns } }
        });
        return files.Select(f => f.Path.LocalPath).ToArray();
    }

    public async Task<string?> OpenFolderAsync(string title)
    {
        var folders = await OpenFoldersInternalAsync(title, false);
        return folders?.FirstOrDefault();
    }

    public async Task<string[]?> OpenFoldersAsync(string title)
    {
        return await OpenFoldersInternalAsync(title, true);
    }

    private async Task<string[]?> OpenFoldersInternalAsync(string title, bool allowMultiple)
    {
        var folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = title,
            AllowMultiple = allowMultiple
        });
        return folders.Select(f => f.Path.LocalPath).ToArray();
    }

    public async Task<string?> SaveFileAsync(string title, string defaultFileName, string extension)
    {
        var ext = extension.TrimStart('.');
        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = title,
            SuggestedFileName = defaultFileName,
            DefaultExtension = ext,
            FileTypeChoices = new[] { new FilePickerFileType($"{ext.ToUpper()} 파일") { Patterns = new[] { $"*.{ext}" } } }
        });
        return file?.Path.LocalPath;
    }
}