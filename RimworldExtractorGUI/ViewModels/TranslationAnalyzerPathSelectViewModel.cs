using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimworldExtractorInternal.DiffA;

namespace RimworldExtractorGUI.ViewModels;

public partial class TranslationAnalyzerPathSelectViewModel : ViewModelBase
{
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CompleteCommand))]
    private string _pathText = string.Empty;

    public string[] Paths { get; private set; } = Array.Empty<string>();
    public bool IsSuccess { get; private set; } = false;

    public event Action? RequestClose;

    [RelayCommand]
    private async Task SelectSingleFileAsync(IStorageProvider storageProvider)
    {
        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "엑셀 파일(.xlsx)을 선택해주세요.",
            AllowMultiple = true,
            FileTypeFilter = new[] { new FilePickerFileType("림왈도 형식 엑셀 파일") { Patterns = new[] { "*.xlsx" } } }
        });

        if (files.Count > 0)
        {
            PathText = string.Join('|', files.Select(f => f.Path.LocalPath));
        }
    }

    [RelayCommand]
    private async Task SelectDirAsync(IStorageProvider storageProvider)
    {
        var folders = await storageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "엑셀 파일이 있는 루트 폴더를 선택해주세요.",
            AllowMultiple = true
        });

        if (folders.Count > 0)
        {
            PathText = string.Join('|', folders.Select(f => f.Path.LocalPath));
        }
    }

    private bool CanComplete() => !string.IsNullOrWhiteSpace(PathText);

    [RelayCommand(CanExecute = nameof(CanComplete))]
    private void Complete()
    {
        var tokens = (PathText ?? "").Split('|', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length > 0 && Directory.Exists(tokens[0]))
        {
            tokens = tokens.SelectMany(DiffAnalyzer.GetXlsxPaths).ToArray();
        }
        Paths = tokens;
        IsSuccess = true;
        RequestClose?.Invoke();
    }
}