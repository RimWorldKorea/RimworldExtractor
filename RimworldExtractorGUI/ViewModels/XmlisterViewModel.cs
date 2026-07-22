using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace RimworldExtractorGUI.ViewModels;

public partial class XmlisterViewModel : ViewModelBase
{
    [ObservableProperty] private string _folderPathText = string.Empty;

    public string[] FileNames { get; private set; } = Array.Empty<string>();
    public bool IsSuccess { get; private set; } = false;

    public event Action? RequestClose;

    [RelayCommand]
    private async Task SelectFoldersAsync(IStorageProvider storageProvider)
    {
        var folders = await storageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Languages 폴더가 있는 루트 폴더를 지정해주세요.",
            AllowMultiple = true
        });

        if (folders.Count > 0)
        {
            FolderPathText = string.Join('|', folders.Select(f => f.Path.LocalPath));
        }
    }

    [RelayCommand]
    private void Complete()
    {
        FileNames = (FolderPathText ?? "").Split('|', StringSplitOptions.RemoveEmptyEntries);
        IsSuccess = true;
        RequestClose?.Invoke();
    }
}