using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimworldExtractorInternal;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorGUI.ViewModels;

public partial class TranslationAnalyzerViewModel : ViewModelBase
{
    [ObservableProperty] private string _titleText = "번역 데이터를 분석하고 있습니다...";
    [ObservableProperty] private string _selectedModTitle = "수정할 모드를 선택하세요.";
    [ObservableProperty] private TranslationAnalyzerItemViewModel? _selectedItem;
    [ObservableProperty] private int _selectedSaveMethodIndex = 0;
    [ObservableProperty] private bool _isControlEnabled = false;

    public ObservableCollection<TranslationAnalyzerItemViewModel> Items { get; } = new();
    public bool IsSuccess { get; private set; } = false;

    public event Func<ModMetadata?, Task<(bool Success, ModMetadata? Mod, List<ExtractableFolder> Folders, List<ModMetadata> RefMods)>>? RequestSelectModDialog;
    public event Func<string, string, Task>? RequestShowAlert;
    public event Action? RequestClose;

    public TranslationAnalyzerViewModel() { }

    public TranslationAnalyzerViewModel(string[] paths)
    {
        Task.Run(() => AnalyzeTranslationAsync(paths));
    }

    partial void OnSelectedItemChanged(TranslationAnalyzerItemViewModel? value)
    {
        if (value == null)
        {
            IsControlEnabled = false;
            SelectedModTitle = "수정할 모드를 선택하세요.";
            return;
        }

        IsControlEnabled = true;
        SelectedModTitle = value.Entry.Metadata?.ToString() ?? "원본 모드를 찾을 수 없었습니다. 수동으로 지정해주세요.";
        SelectedSaveMethodIndex = (int)value.Entry.SaveMethod;
    }

    partial void OnSelectedSaveMethodIndexChanged(int value)
    {
        if (SelectedItem == null || value < 0) return;

        SelectedItem.Entry.SaveMethod = (TranslationAnalyzerEntry.SaveMethodEnum)value;
        SelectedItem.UpdateSaveMethodText();
    }

    private async Task AnalyzeTranslationAsync(string[] paths)
    {
        int invalidCount = 0;
        for (var i = 0; i < paths.Length; i++)
        {
            var currentIdx = i;
            var path = paths[i];

            Dispatcher.UIThread.Post(() =>
            {
                TitleText = $"번역 데이터를 분석하고 있습니다... {currentIdx}/{paths.Length}";
            });

            var entry = new TranslationAnalyzerEntry(path);
            if (entry.Metadata != null)
            {
                var autoSelectedFolders = ModLister.GetExtractableFolders(entry.Metadata)
                    .Where(x => x.IsAutoSelectable()).ToList();

                var autoSelectedRefMods = ModLister.FindAllReferenceMods(entry.Metadata).Distinct().ToList();
                entry.ReExtract(autoSelectedFolders, autoSelectedRefMods);
            }

            if (entry.Invalid) invalidCount++;

            var itemVm = new TranslationAnalyzerItemViewModel(entry);

            Dispatcher.UIThread.Post(() =>
            {
                Items.Add(itemVm);
            });
        }

        Dispatcher.UIThread.Post(async () =>
        {
            TitleText = "분석 완료!";
            if (invalidCount > 0 && RequestShowAlert != null)
            {
                await RequestShowAlert.Invoke("경고", "몇몇 엑셀 파일들을 정상적으로 분석하지 못했습니다. 로그창 확인 후 다시 시도해주세요.");
            }
        });
    }

    [RelayCommand]
    private void SelectAll()
    {
        foreach (var item in Items)
        {
            if (item.Entry.Metadata != null && item.Entry.HasChanges)
                item.IsChecked = true;
        }
    }

    [RelayCommand]
    private void DeselectAll()
    {
        foreach (var item in Items)
        {
            item.IsChecked = false;
        }
    }

    [RelayCommand]
    private async Task OpenSelectModAsync()
    {
        if (SelectedItem == null || RequestSelectModDialog == null) return;

        var result = await RequestSelectModDialog.Invoke(SelectedItem.Entry.Metadata);
        if (result.Success && result.Mod != null)
        {
            SelectedItem.Entry.Metadata = result.Mod;
            SelectedItem.Entry.ResetChanges();
            SelectedItem.Entry.ReExtract(result.Folders, result.RefMods);
            SelectedItem.UpdateFromEntry();

            SelectedModTitle = SelectedItem.Entry.Metadata.ToString();
        }
    }

    [RelayCommand]
    private void Apply()
    {
        IsSuccess = true;
        RequestClose?.Invoke();
    }

    public IEnumerable<TranslationAnalyzerEntry> GetSelectedEntries()
    {
        return Items.Where(x => x.IsChecked && x.Entry.HasChanges)
                    .Select(x => x.Entry);
    }
}