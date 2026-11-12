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
    public Mod? Metadata { get; }
    
    /// 구분선 용도로 쓰는 껍데기 개체 여부입니다.
    public bool IsHeader { get; }
    public string HeaderText { get; } = string.Empty;

    public string DisplayId { get; } = string.Empty;
    public string DisplayName { get; } = string.Empty;

    public double MinIdWidth => _parentVm!.IdColumnMinWidth;

    [ObservableProperty]
    private bool _isReference;
    
    [ObservableProperty] 
    private bool _isExtractionTarget;

    // 1. 헤더용 생성자
    public ModListItem(string headerText)
    {
        IsHeader = true;
        HeaderText = headerText;
    }

    // 2. 일반 아이템용 생성자
    public ModListItem(Mod metadata, bool isReference, SelectModViewModel parentVm)
    {
        IsHeader = false;
        Metadata = metadata;
        _isReference = isReference;
        _parentVm = parentVm;

        DisplayId = metadata.Type == ModType.Official ||
                    string.IsNullOrWhiteSpace(metadata.WorkshopID) ||
                    metadata.WorkshopID == "???" 
            ? "-" 
            : metadata.WorkshopID;
            
        DisplayName = metadata.Name ?? "Unknown";
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
        
        _parentVm.UpdateSelectedModInfoText();
    }

    public override string ToString() => IsHeader ? HeaderText : DisplayName;
}

public partial class SelectModViewModel : ViewModelBase
{
    [ObservableProperty] private bool _isLoading = true;
    [ObservableProperty] private double _idColumnMinWidth = 40.0; // 뷰(UI)에 바인딩할 동적 최소 너비 속성
    
    private readonly IExternalProcessService _processService;
    private readonly IExtractionService _extractionService;
    
    private List<Mod> _allModsCached = new();
    private List<Mod> _officialModsCached = new();
    private List<Mod> _localModsCached = new();
    private List<Mod> _workshopModsCached = new();
    private readonly Mod? _initialFocusingMod;
    private readonly List<Mod>? _initialReferenceMods;

    /// <todo>이게 여기 있는 것이 그다지 적절하지 않을 것이다.</todo>
    private static readonly List<string> OfficialDLCSequence = new()
    { "Core", "Royalty", "Ideology", "Biotech", "Anomaly", "Odessey" };

    [ObservableProperty] private string _searchText = string.Empty;
    [ObservableProperty] private bool _isFilterSelectedOnly = false;
    [ObservableProperty] private string _selectedModInfoText = "선택된 모드가 없습니다.";
    [ObservableProperty] private bool _autoReferenceLink;
    
    /// <summary>SelectedMod의 수정 요청을 임시로 받아두기 위한 속성입니다.</summary>
    [ObservableProperty] private ModListItem? _selectedModListItem;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(CompleteCommand))]
    private bool _canComplete = false;
    
    public Mod? SelectedMod { get; private set; }
    public List<Mod> ReferenceMods { get; } = new();
    
    public ObservableCollection<ModListItem> FilteredMods { get; } = new();

    public event Action? RequestClose;
    public event Func<string, Task>? RequestShowAlert;

    public SelectModViewModel(IExternalProcessService processService, IExtractionService extractionService, Mod? initialFocusingMod = null, List<Mod>? initialReferenceMods = null)
    {
        _processService = processService;
        _extractionService = extractionService;
        _initialFocusingMod = initialFocusingMod;
        _initialReferenceMods = initialReferenceMods;
        
        AutoReferenceLink = SettingManager.Current.AutoReferenceLink;
    }

    /// <summary>UI와 분리된 비동기 초기화 메서드입니다.</summary>
    public async Task InitializeAsync()
    {
        IsLoading = true;
        int maxIdLength = 0;

        await Task.Run(() =>
        {
            ModLister.ResetCache();
        
            _officialModsCached = ModLister.OfficialMods.OrderBy(m => {
                int index = OfficialDLCSequence.IndexOf(m.Name ?? "");
                return index == -1 ? int.MaxValue : index;
            }).ToList();
        
            _localModsCached = ModLister.LocalMods.ToList();
            _workshopModsCached = ModLister.WorkshopMods.ToList();
            _allModsCached = _officialModsCached.Concat(_localModsCached).Concat(_workshopModsCached).ToList();

            foreach (var mod in _allModsCached)
            {
                string id = mod.Type == ModType.Official || string.IsNullOrWhiteSpace(mod.WorkshopID) || mod.WorkshopID == "???" 
                    ? "-" : mod.WorkshopID;
                if (id.Length > maxIdLength) maxIdLength = id.Length;
            }
        });

        // UI 스레드 복귀 후 UI 바인딩 갱신
        IdColumnMinWidth = Math.Max(40.0, (maxIdLength * 8) + 26);

        if (_initialReferenceMods != null && _initialReferenceMods.Count > 0)
        {
            var initialPaths = _initialReferenceMods.Select(m => m.Path).ToHashSet();
            foreach (var mod in _allModsCached)
                if (initialPaths.Contains(mod.Path))
                    ReferenceMods.Add(mod);
        }
        else
        {
            foreach (var officialMod in _officialModsCached)
                if (!ReferenceMods.Contains(officialMod))
                    ReferenceMods.Add(officialMod);
        }
        
        if (_initialFocusingMod != null)
        {
            SelectedMod = _allModsCached.FirstOrDefault(m
                => m.Path == _initialFocusingMod.Path);
            
            if (SelectedMod != null) CanComplete = true;
        }
        
        RefreshModList();
        
        if (SelectedMod != null)
        {
            SelectedModListItem = FilteredMods.FirstOrDefault(x
                => !x.IsHeader && x.Metadata?.Path == SelectedMod.Path);
        }
        
        UpdateSelectedModInfoText();
        
        IsLoading = false;
    }
    
    partial void OnAutoReferenceLinkChanged(bool value)
    {
        SettingManager.Current.AutoReferenceLink = value;
        SettingManager.LastSaved.AutoReferenceLink = value;
        SettingManager.SaveToDisk();
    }
    partial void OnSearchTextChanged(string value) => RefreshModList();
    partial void OnIsFilterSelectedOnlyChanged(bool value) => RefreshModList();
    partial void OnSelectedModListItemChanged(ModListItem? value)
    {
        if (_isLoading) return;
        // 헤더 선택시 무시
        if (value == null || value.IsHeader || value.Metadata == null) return;
  
        if (AutoReferenceLink) ResetSelection(false);
        SelectedMod = value.Metadata;
        if (AutoReferenceLink) _ = SelectAllRefModsInternalAsync(value, false);
        
        RefreshModList();
        UpdateSelectedModInfoText();
        CanComplete = true;
    }

    /// <summary>정보창 텍스트를 갱신합니다.</summary>
    public void UpdateSelectedModInfoText()
    {
        if (SelectedMod == null)
        {
            SelectedModInfoText = "선택된 모드가 없습니다.";
            return;
        }

        var info = "추출 모드: " + SelectedMod.Name;
        if (ReferenceMods.Count > 0)
        {
            info += "\n참조 모드: ";
            info += string.Join(", ", ReferenceMods.Select(m => m.Name)); 
        }
        SelectedModInfoText = info;
    }
    
    /// <summary>모드 리스트 UI를 갱신합니다.</summary>
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

    private IEnumerable<ModListItem> GetFilteredMods(IEnumerable<Mod> mods, string keyword)
    {
        foreach (var mod in mods)
        {
            var identifier = mod.Type == ModType.Official ? mod.Name : $"{mod.Name} - {mod.WorkshopID}";
            if (string.IsNullOrEmpty(keyword) || (identifier?.ToLower().Contains(keyword) ?? false))
            {
                bool isRef = ReferenceMods.Contains(mod);
                if (IsFilterSelectedOnly && !isRef && SelectedMod != mod)
                    continue;

                var item = new ModListItem(mod, isRef, this);
                
                // 방금 만든 아이템이 현재 타겟 모드(SelectedMod)라면 true
                if (SelectedMod != null && mod.Path == SelectedMod.Path)
                    item.IsExtractionTarget = true;
                
                yield return item;
            }
        }
    }

    [RelayCommand]
    private async Task SelectAllRefModsAsync(ModListItem? clickedItem)
    {
        await SelectAllRefModsInternalAsync(clickedItem ?? SelectedModListItem, true);
    }
    
    /// <summary>선택한 모드의 참조모드 목록을 가져오는 메서드.</summary>
    /// <todo>뷰모델이 아니라 ModLister로 로직을 옮기는게 맞는거 아닌가?</todo>
    private async Task SelectAllRefModsInternalAsync(ModListItem? targetItem, bool refreshUI)
    {
        if (targetItem?.Metadata == null) return;
        var mod = targetItem.Metadata;
        
        // 1. 기본 종속성과 버전별 종속성을 모두 담을 집합(중복 방지)
        var allDependencies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        
        foreach (var pid in mod.About.modDependencies) 
            allDependencies.Add(pid);
        
        // 2. 현재 버전에 맞는 종속성(modDependenciesByVersion) 추가
        if (Version.TryParse(SettingManager.Current.CurrentVersion, out var currentVersion))
        {
            if (mod.About?.modDependenciesByVersion != null && 
                mod.About.modDependenciesByVersion.TryGetValue(currentVersion, out var versionDeps))
            {
                foreach (var pid in versionDeps)
                    allDependencies.Add(pid);
            }
        }

        // 3. 수집된 종속성 ID들을 캐시된 모드 리스트에서 검색
        var requiredMods = allDependencies
            .Select(x => _allModsCached.Find(y => string.Equals(y.PackageID, x, StringComparison.OrdinalIgnoreCase)))
            .Where(x => x != null)
            .ToList();
        
        // 실제 참조 모드 로드
        foreach (var refMod in ModLister.FindAllReferenceMods(mod))
        {
            if (!ReferenceMods.Contains(refMod))
                ReferenceMods.Add(refMod);
        }

        if (refreshUI)
        {
            RefreshModList();
            UpdateSelectedModInfoText();
        }
    }

    [RelayCommand]
    private async Task OpenExplorerAsync(ModListItem? clickedItem)
    {
        var targetItem = clickedItem ?? SelectedModListItem;
        
        if (targetItem?.Metadata?.Path != null)
        {
            var path = targetItem.Metadata.Path;
            var targetMod = targetItem.Metadata;
            
            // 폴더 존재 여부를 뷰모델 단계에서 먼저 확인
            if (!System.IO.Directory.Exists(path))
            {
                if (RequestShowAlert != null)
                    await RequestShowAlert.Invoke("존재하지 않는 폴더입니다.");
                
                // 1. 뷰모델의 캐시 리스트에서 해당 모드 제거
                _allModsCached.Remove(targetMod);
                _officialModsCached.Remove(targetMod);
                _localModsCached.Remove(targetMod);
                _workshopModsCached.Remove(targetMod);
                
                // 2. 참조 모드로 선택되어 있었다면 제거
                if (ReferenceMods.Contains(targetMod))
                    ReferenceMods.Remove(targetMod);

                // 3. 현재 선택된 모드 초기화
                if (SelectedModListItem == targetItem)
                {
                    SelectedModListItem = null;
                    SelectedModInfoText = "선택된 모드가 없습니다.";
                    CanComplete = false;
                }

                // 4. Core 엔진 레지스트리(RegisteredMods)에서 영구 삭제 및 자원 해제
                targetMod.Dispose();
                
                // 5. 남은 캐시를 바탕으로 UI 리스트 갱신
                RefreshModList(); 
                return;
            }

            _processService.OpenFolderInExplorer(path);
        }
    }
    
    /// <summary>모든 선택 상태를 초기 상태로 되돌립니다.</summary>
    [RelayCommand]
    private void ResetSelection(bool refreshUI = true)
    {
        SelectedMod = null;
        ReferenceMods.Clear();
        SearchText = string.Empty;
        IsFilterSelectedOnly = false;
        
        foreach (var officialMod in _officialModsCached)
        {
            if (!ReferenceMods.Contains(officialMod))
                ReferenceMods.Add(officialMod);
        }
        
        if(refreshUI)
        {
            RefreshModList();
            UpdateSelectedModInfoText();
        }
    }

    private bool CanExecuteComplete() => CanComplete && SelectedMod != null;
    [RelayCommand(CanExecute = nameof(CanExecuteComplete))]
    private void Complete()
    {
        RequestClose?.Invoke();
    }
}