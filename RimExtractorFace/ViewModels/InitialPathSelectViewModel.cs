using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimExtractorCore;

namespace RimExtractorFace.ViewModels;

public partial class InitialPathSelectViewModel : ViewModelBase
{
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CompleteCommand))]
    private string _pathRimworld = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CompleteCommand))]
    private string _pathWorkshop = string.Empty;

    public event Action? RequestClose;

    public InitialPathSelectViewModel()
    {
        ConfigManager.InitDefault();
        PathRimworld = ConfigManager.Current.PathRimworld;
        PathWorkshop = ConfigManager.Current.PathWorkshop;
    }

    // 림월드 실행 파일 선택 명령
    [RelayCommand]
    private async Task SelectPathRimworldAsync(IStorageProvider storageProvider)
    {
        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "RimWorldWin64.exe를 지정해주세요",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("림월드 실행 파일") { Patterns = new[] { "RimWorldWin64.exe" } }
            }
        });

        if (files.Count > 0)
        {
            PathRimworld = Path.GetDirectoryName(files[0].Path.LocalPath) ?? string.Empty;
        }
    }

    // 창작마당 폴더 선택 명령
    [RelayCommand]
    private async Task SelectPathWorkshopAsync(IStorageProvider storageProvider)
    {
        var folders = await storageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "림월드 창작마당 경로를 지정해주세요 => Steam\\steamapps\\workshop\\content\\294100",
            AllowMultiple = false
        });

        if (folders.Count > 0)
        {
            PathWorkshop = folders[0].Path.LocalPath;
        }
    }

    private bool CanComplete() => 
        !string.IsNullOrWhiteSpace(PathRimworld) && !string.IsNullOrWhiteSpace(PathWorkshop);

    // 완료 명령
    [RelayCommand(CanExecute = nameof(CanComplete))]
    private void Complete()
    {
        ConfigManager.Current.PathRimworld = PathRimworld;
        ConfigManager.Current.PathWorkshop = PathWorkshop;
        ConfigManager.Save();

        RequestClose?.Invoke();
    }
}