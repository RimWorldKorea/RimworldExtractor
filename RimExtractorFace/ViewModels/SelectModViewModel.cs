using System.Collections.ObjectModel;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimExtractorCore;
using RimExtractorCore.DataTypes;
using RimExtractorFace.Services;

namespace RimExtractorFace.ViewModels;

public partial class ModListItem : ObservableObject
{
    private readonly SelectModViewModel? _parentVm;
    public ModMetadata? Metadata { get; }
    
    // [복구됨] 헤더인지 일반 아이템인지 구분
    public bool IsHeader { get; }
    public string HeaderText { get; } = string.Empty;

    public string DisplayId { get; } = string.Empty;
    public string DisplayName { get; } = string.Empty;

    public double MinIdWidth => _parentVm!.IdColumnMinWidth;

    [ObservableProperty]
    private bool _isReference;

    // 1. 헤더용 생성자
    public ModListItem(string headerText)
    {
        IsHeader = true;
        HeaderText = headerText;
    }

    // 2. 일반 아이템용 생성자
    public ModListItem(ModMetadata metadata, bool isReference, SelectModViewModel parentVm)
    {
        IsHeader = false;
        Metadata = metadata;
        _isReference = isReference;
        _parentVm = parentVm;
        
        //TODO "???" 크아악
        DisplayId = metadata.IsOfficialContent || string.IsNullOrWhiteSpace(metadata.Id) || metadata.Id == "???" 
            ? "-" 
            : metadata.Id;
            
        DisplayName = metadata.ModName;
    }

    partial void OnIsReferenceChanged(bool value)
    {
        if (_parentVm == null || Metadata == null) return;
        
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

    public override string ToString() => IsHeader ? HeaderText : DisplayName;
}

public partial class SelectModViewModel : ViewModelBase
{
    private readonly IExternalProcessService _processService;
    private readonly IExtractionService _extractionService;
    private readonly List<ModMetadata> _allModsCached;
    private readonly List<ModMetadata> _officialModsCached;
    private readonly List<ModMetadata> _localModsCached;
    private readonly List<ModMetadata> _workshopModsCached;

    private static readonly List<string> OfficialDLCSequence = new()
    {
        "Core", "Royalty", "Ideology", "Biotech", "Anomaly", "Odessey"
    };

    [ObservableProperty] private string _searchText = string.Empty;
    [ObservableProperty] private bool _isFilterSelectedOnly = false;
    [ObservableProperty] private string _selectedModInfoText = "선택된 모드가 없습니다.";
    
    [ObservableProperty] private ModListItem? _selectedModListItem;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(CompleteCommand))]
    private bool _canComplete = false;
    
    // 뷰(UI)에 바인딩할 동적 최소 너비 속성
    public double IdColumnMinWidth { get; }

    public ModMetadata? SelectedMod { get; private set; }
    public List<ExtractableFolder> SelectedFolders { get; } = new();
    public List<ModMetadata> ReferenceMods { get; } = new();
    
    // [복구됨] 다시 단일 리스트 하나만 사용합니다.
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
        
        _officialModsCached = ModLister.OfficialMods
            .OrderBy(m => 
            {
                int index = OfficialDLCSequence.IndexOf(m.ModName);
                return index == -1 ? int.MaxValue : index;
            }).ToList();
            
        _localModsCached = ModLister.LocalMods.ToList();
        _workshopModsCached = ModLister.WorkshopMods.ToList();
        _allModsCached = _officialModsCached.Concat(_localModsCached).Concat(_workshopModsCached).ToList();

        // [추가됨] 창이 열릴 때 한 번만 전체 모드를 스캔하여 ID 최대 길이를 계산합니다.
        int maxIdLength = 0;
        foreach (var mod in _allModsCached)
        {
            //TODO "???"은 어디서 달라붙는겨
            string id = mod.IsOfficialContent || string.IsNullOrWhiteSpace(mod.Id) || mod.Id == "???" ? "-" : mod.Id;
            if (id.Length > maxIdLength) 
                maxIdLength = id.Length;
        }
        // 폰트 크기를 고려해 '글자 수 * 약 8px'로 계산하고 좌우 마진(16) + 여유(10)를 더해줍니다. 기본 하한선은 70입니다.
        IdColumnMinWidth = Math.Max(40.0, (maxIdLength * 8) + 26);
        
        
        foreach (var officialMod in _officialModsCached)
        {
            if (!ReferenceMods.Contains(officialMod))
                ReferenceMods.Add(officialMod);
        }

        RefreshModList();

        if (initialMod != null)
        {
            SelectedModListItem = FilteredMods.FirstOrDefault(x => !x.IsHeader && x.Metadata?.RootDir == initialMod.RootDir);
        }
    }

    partial void OnSearchTextChanged(string value) => RefreshModList();
    partial void OnIsFilterSelectedOnlyChanged(bool value) => RefreshModList();
    partial void OnSelectedModListItemChanged(ModListItem? value)
    {
        // [예외 처리] 헤더(가짜 아이템)를 클릭했을 때는 무시합니다.
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
        var folders = ModLister.GetExtractableFolders(SelectedMod, SettingManager.Current.CurrentVersion);
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

        var official = GetFilteredMods(_officialModsCached, keyword).ToList();
        if (official.Any())
        {
            FilteredMods.Add(new ModListItem("림월드 공식 컨텐츠"));
            foreach (var mod in official) FilteredMods.Add(mod);
        }

        var local = GetFilteredMods(_localModsCached, keyword).ToList();
        if (local.Any())
        {
            FilteredMods.Add(new ModListItem("로컬 모드"));
            foreach (var mod in local) FilteredMods.Add(mod);
        }

        var workshop = GetFilteredMods(_workshopModsCached, keyword).ToList();
        if (workshop.Any())
        {
            FilteredMods.Add(new ModListItem("창작마당 모드"));
            foreach (var mod in workshop) FilteredMods.Add(mod);
        }
    }

    private IEnumerable<ModListItem> GetFilteredMods(IEnumerable<ModMetadata> mods, string keyword)
    {
        foreach (var mod in mods)
        {
            if (string.IsNullOrEmpty(keyword) || mod.Identifier.ToLower().Contains(keyword))
            {
                bool isRef = ReferenceMods.Contains(mod);
                if (IsFilterSelectedOnly && !isRef && SelectedMod != mod)
                    continue;

                yield return new ModListItem(mod, isRef, this);
            }
        }
    }

    [RelayCommand]
    private async Task SelectAllRefModsAsync()
    {
        if (SelectedModListItem?.Metadata == null) return;
        var mod = SelectedModListItem.Metadata;
        
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
        
        RefreshModList();
    }

    [RelayCommand]
    private void OpenExplorer()
    {
        if (SelectedModListItem?.Metadata != null)
        {
            _processService.OpenFolderInExplorer(SelectedModListItem.Metadata.RootDir);
        }
    }

    private bool CanExecuteComplete() => CanComplete && SelectedMod != null;
    [RelayCommand(CanExecute = nameof(CanExecuteComplete))]
    private void Complete()
    {
        RequestClose?.Invoke();
    }
}