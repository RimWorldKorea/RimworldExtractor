using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimworldExtractorInternal;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorGUI.ViewModels;

public class ModListItem
{
    public ModMetadata? Metadata { get; }
    public string DisplayText { get; }
    public bool IsHeader { get; }

    public ModListItem(string header)
    {
        DisplayText = header;
        IsHeader = true;
    }

    public ModListItem(ModMetadata metadata, bool isReference)
    {
        Metadata = metadata;
        IsHeader = false;
        
        string prefix = isReference ? "(참조) " : "";
        DisplayText = metadata.IsOfficialContent
            ? $"{prefix}[Official] {metadata.ModName}"
            : $"{prefix}[{metadata.Id}] {metadata.ModName}";
    }

    public override string ToString() => DisplayText;
}

public partial class SelectModViewModel : ViewModelBase
{
    private readonly List<ModMetadata> _allModsCached;
    private readonly List<ModMetadata> _officialModsCached;
    private readonly List<ModMetadata> _localModsCached;
    private readonly List<ModMetadata> _workshopModsCached;

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private bool _isFilterSelectedOnly = false;

    [ObservableProperty]
    private string _selectedModInfoText = "추출할 모드를 선택하세요";

    [ObservableProperty]
    private ModListItem? _selectedModListItem;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CompleteCommand))]
    private bool _canComplete = false;

    public ModMetadata? SelectedMod { get; private set; }
    public List<ExtractableFolder> SelectedFolders { get; } = new();
    public List<ModMetadata> ReferenceMods { get; } = new();

    public ObservableCollection<ModListItem> FilteredMods { get; } = new();
    public ObservableCollection<ExtractableFolder> ExtractableFolders { get; } = new();

    public event Action? RequestClose;
    public event Func<string, Task>? RequestShowAlert;

    public SelectModViewModel(ModMetadata? initialMod = null)
    {
        ModLister.ResetCache();
        _officialModsCached = ModLister.OfficialMods.ToList();
        _localModsCached = ModLister.LocalMods.ToList();
        _workshopModsCached = ModLister.WorkshopMods.ToList();
        _allModsCached = _officialModsCached.Concat(_localModsCached).Concat(_workshopModsCached).ToList();

        // 기본 참조 모드 리스트 로드
        if (!string.IsNullOrEmpty(Prefabs.PathBaseRefList) && File.Exists(Prefabs.PathBaseRefList))
        {
            var lines = File.ReadAllLines(Prefabs.PathBaseRefList);
            foreach (var mod in _allModsCached)
            {
                if (lines.Any(x => mod.Identifier == x))
                {
                    ReferenceMods.Add(mod);
                }
            }
        }

        RefreshModList();

        if (initialMod != null)
        {
            SelectedModListItem = FilteredMods.FirstOrDefault(x => x.Metadata?.RootDir == initialMod.RootDir);
        }
    }

    partial void OnSearchTextChanged(string value) => RefreshModList();
    partial void OnIsFilterSelectedOnlyChanged(bool value) => RefreshModList();

    partial void OnSelectedModListItemChanged(ModListItem? value)
    {
        if (value == null || value.IsHeader || value.Metadata == null)
            return;

        SelectedMod = value.Metadata;
        
        var info = SelectedMod.ModName;
        if (SelectedMod.ModDependencies is { Count: > 0 })
        {
            info += $"\n[선행모드: {string.Join(';', SelectedMod.ModDependencies)}]";
        }
        SelectedModInfoText = info;

        ExtractableFolders.Clear();
        var folders = ModLister.GetExtractableFolders(SelectedMod);
        foreach (var folder in folders)
        {
            ExtractableFolders.Add(folder);
        }

        CanComplete = ExtractableFolders.Count > 0;
    }

    public void RefreshModList()
    {
        FilteredMods.Clear();
        var keyword = SearchText.Trim().ToLower();

        // Official
        FilteredMods.Add(new ModListItem("==================== OFFICIAL ===================="));
        AddFilteredItems(_officialModsCached, keyword);

        // Local
        FilteredMods.Add(new ModListItem("==================== LOCAL MODS ===================="));
        AddFilteredItems(_localModsCached, keyword);

        // Workshop
        FilteredMods.Add(new ModListItem("==================== WORKSHOP MODS ===================="));
        AddFilteredItems(_workshopModsCached, keyword);
    }

    private void AddFilteredItems(IEnumerable<ModMetadata> mods, string keyword)
    {
        foreach (var mod in mods)
        {
            if (string.IsNullOrEmpty(keyword) || mod.Identifier.ToLower().Contains(keyword))
            {
                bool isRef = ReferenceMods.Contains(mod);
                if (IsFilterSelectedOnly && !isRef && SelectedMod != mod)
                    continue;

                FilteredMods.Add(new ModListItem(mod, isRef));
            }
        }
    }

    [RelayCommand]
    private void ToggleRefMod()
    {
        if (SelectedModListItem?.Metadata is not { } mod) return;

        if (ReferenceMods.Contains(mod))
            ReferenceMods.Remove(mod);
        else
            ReferenceMods.Add(mod);

        RefreshModList();
    }

    [RelayCommand]
    private async Task SelectAllRefModsAsync()
    {
        if (SelectedModListItem?.Metadata is not { } mod) return;

        var requiredMods = mod.ModDependencies?.Select(x => _allModsCached.Find(y => y.PackageId == x)).ToList();
        if (requiredMods == null || requiredMods.Count == 0)
        {
            if (RequestShowAlert != null)
                await RequestShowAlert.Invoke("이 모드는 선행 모드나 선택적 선행 모드가 없습니다!");
            return;
        }

        foreach (var refMod in ModLister.FindAllReferenceMods(mod))
        {
            if (!ReferenceMods.Contains(refMod))
                ReferenceMods.Add(refMod);
        }

        RefreshModList();
    }

    [RelayCommand]
    private void OpenExplorer()
    {
        if (SelectedModListItem?.Metadata is { } mod)
        {
            Process.Start(new ProcessStartInfo { FileName = mod.RootDir, UseShellExecute = true });
        }
    }

    [RelayCommand]
    private async Task SaveRefModsListAsync(IStorageProvider storageProvider)
    {
        var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "선택한 참조 모드의 목록을 저장합니다.",
            DefaultExtension = "refMods",
            FileTypeChoices = new[] { new FilePickerFileType("참조 모드 리스트 파일") { Patterns = new[] { "*.refMods" } } }
        });

        if (file != null)
        {
            await File.WriteAllLinesAsync(file.Path.LocalPath, ReferenceMods.Select(x => x.Identifier));
        }
    }

    [RelayCommand]
    private async Task LoadRefModsListAsync(IStorageProvider storageProvider)
    {
        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "선택한 파일로부터 참조 모드의 목록을 불러옵니다.",
            AllowMultiple = false,
            FileTypeFilter = new[] { new FilePickerFileType("참조 모드 리스트 파일") { Patterns = new[] { "*.refMods" } } }
        });

        if (files.Count > 0)
        {
            ReferenceMods.Clear();
            var lines = await File.ReadAllLinesAsync(files[0].Path.LocalPath);
            foreach (var mod in _allModsCached)
            {
                if (lines.Any(x => mod.Identifier == x))
                    ReferenceMods.Add(mod);
            }
            RefreshModList();
        }
    }

    private bool CanExecuteComplete() => CanComplete && SelectedMod != null;

    [RelayCommand(CanExecute = nameof(CanExecuteComplete))]
    private void Complete()
    {
        RequestClose?.Invoke();
    }
}