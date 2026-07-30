// ViewModels/SelectModViewModel.cs 전체 코드를 아래로 교체해주세요.

using System.Collections.ObjectModel;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimExtractorCore;
using RimExtractorCore.DataTypes;
using RimExtractorFace.Services;

namespace RimExtractorFace.ViewModels;

// [수정됨] 체크박스 상태 변경을 감지하기 위해 ObservableObject 상속 및 partial 클래스로 변경
public partial class ModListItem : ObservableObject
{
    private readonly SelectModViewModel? _parentVm;
    public ModMetadata? Metadata { get; }
    public string DisplayText { get; }
    public bool IsHeader { get; }

    [ObservableProperty]
    private bool _isReference;

    public ModListItem(string header)
    {
        DisplayText = header;
        IsHeader = true;
    }

    public ModListItem(ModMetadata metadata, bool isReference, SelectModViewModel parentVm)
    {
        Metadata = metadata;
        IsHeader = false;
        _isReference = isReference;
        _parentVm = parentVm;
        
        // (참조) 접두사는 체크박스로 대체하므로 제거
        DisplayText = metadata.IsOfficialContent
            ? $"[Official] {metadata.ModName}"
            : $"[{metadata.Id}] {metadata.ModName}";
    }

    // 체크박스가 눌려 상태가 변할 때 리스트에 넣고 빼기
    partial void OnIsReferenceChanged(bool value)
    {
        if (Metadata == null || _parentVm == null) return;
        
        if (value)
        {
            if (!_parentVm.ReferenceMods.Contains(Metadata))
                _parentVm.ReferenceMods.Add(Metadata);
        }
        else
        {
            _parentVm.ReferenceMods.Remove(Metadata);
        }
    }

    public override string ToString() => DisplayText;
}

public partial class SelectModViewModel : ViewModelBase
{
    private readonly IExternalProcessService _processService;
    private readonly IExtractionService _extractionService;
    private readonly List<ModMetadata> _allModsCached;
    private readonly List<ModMetadata> _officialModsCached;
    private readonly List<ModMetadata> _localModsCached;
    private readonly List<ModMetadata> _workshopModsCached;

    [ObservableProperty] private string _searchText = string.Empty;
    [ObservableProperty] private bool _isFilterSelectedOnly = false;
    [ObservableProperty] private string _selectedModInfoText = "선택된 모드가 없습니다.";
    [ObservableProperty] private ModListItem? _selectedModListItem;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(CompleteCommand))]
    private bool _canComplete = false;

    public ModMetadata? SelectedMod { get; private set; }
    public List<ExtractableFolder> SelectedFolders { get; } = new();
    public List<ModMetadata> ReferenceMods { get; } = new();

    public ObservableCollection<ModListItem> FilteredMods { get; } = new();
    public ObservableCollection<ExtractableFolder> ExtractableFolders { get; } = new();

    public event Action? RequestClose;
    public event Func<string, Task>? RequestShowAlert;

    public SelectModViewModel(
        IExternalProcessService processService,
        IExtractionService extractionService,
        ModMetadata? initialMod = null)
    {
        _processService = processService;
        _extractionService = extractionService;
        
        ModLister.ResetCache();
        _officialModsCached = ModLister.OfficialMods.ToList();
        _localModsCached = ModLister.LocalMods.ToList();
        _workshopModsCached = ModLister.WorkshopMods.ToList();
        _allModsCached = _officialModsCached.Concat(_localModsCached).Concat(_workshopModsCached).ToList();

        // [수정됨] 더 이상 refMods 파일을 읽지 않고, 공식(Official) 모드들을 무조건 기본 기준 모드로 추가합니다.
        foreach (var officialMod in _officialModsCached)
        {
            if (!ReferenceMods.Contains(officialMod))
            {
                ReferenceMods.Add(officialMod);
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
            info += $"\n[종속성: {string.Join(';', SelectedMod.ModDependencies)}]";
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

        FilteredMods.Add(new ModListItem("==================== OFFICIAL ===================="));
        AddFilteredItems(_officialModsCached, keyword);

        FilteredMods.Add(new ModListItem("==================== LOCAL MODS ===================="));
        AddFilteredItems(_localModsCached, keyword);

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
                // [수정됨] ViewModel 자신을 넘겨주어 체크박스 이벤트를 처리할 수 있게 합니다.
                FilteredMods.Add(new ModListItem(mod, isRef, this));
            }
        }
    }

    [RelayCommand]
    private async Task SelectAllRefModsAsync()
    {
        if (SelectedModListItem?.Metadata is not { } mod) return;
        
        var requiredMods = mod.ModDependencies?.Select(x => _allModsCached.Find(y => y.PackageId == x)).ToList();
        if (requiredMods == null || requiredMods.Count == 0)
        {
            if (RequestShowAlert != null)
                await RequestShowAlert.Invoke("종속성이 선언되지 않은 모드입니다!");
            return;
        }
        
        foreach (var refMod in ModLister.FindAllReferenceMods(mod))
        {
            if (!ReferenceMods.Contains(refMod))
                ReferenceMods.Add(refMod);
        }
        
        // 추가 후 UI 체크박스 상태 동기화를 위해 새로고침
        RefreshModList();
    }

    [RelayCommand]
    private void OpenExplorer()
    {
        if (SelectedModListItem?.Metadata is { } mod)
        {
            _processService.OpenFolderInExplorer(mod.RootDir);
        }
    }

    private bool CanExecuteComplete() => CanComplete && SelectedMod != null;
    [RelayCommand(CanExecute = nameof(CanExecuteComplete))]
    private void Complete()
    {
        RequestClose?.Invoke();
    }
}