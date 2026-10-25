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
    private readonly Mod? _initialMod;

    //TODO Core 모듈 Constants로 이동
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
    
    public Mod? SelectedMod { get; private set; }
    public List<Mod> ReferenceMods { get; } = new();
    
    public ObservableCollection<ModListItem> FilteredMods { get; } = new();

    public event Action? RequestClose;
    public event Func<string, Task>? RequestShowAlert;

    public SelectModViewModel(IExternalProcessService processService, IExtractionService extractionService, Mod? initialMod = null)
    {
        _processService = processService;
        _extractionService = extractionService;
        _initialMod = initialMod;
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

        foreach (var officialMod in _officialModsCached)
        {
            if (!ReferenceMods.Contains(officialMod))
                ReferenceMods.Add(officialMod);
        }

        RefreshModList();

        if (_initialMod != null)
        {
            SelectedModListItem = FilteredMods.FirstOrDefault(x => !x.IsHeader && x.Metadata?.Path == _initialMod.Path);
        }

        IsLoading = false;
    }
    
    partial void OnSearchTextChanged(string value) => RefreshModList();
    partial void OnIsFilterSelectedOnlyChanged(bool value) => RefreshModList();
    partial void OnSelectedModListItemChanged(ModListItem? value)
    {
        // [예외 처리] 헤더(가짜 아이템)를 클릭했을 때는 무시합니다.
        if (value == null || value.IsHeader || value.Metadata == null) 
            return;
            
        SelectedMod = value.Metadata;
        var info = SelectedMod.Name;
        if (SelectedMod.About?.modDependencies is { Count: > 0 })
        {
            info += $"\n[종속성: {string.Join(';', SelectedMod.About.modDependencies)}]";
        }
        SelectedModInfoText = info;
        
        CanComplete = true;
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

                yield return new ModListItem(mod, isRef, this);
            }
        }
    }

    [RelayCommand]
    private async Task SelectAllRefModsAsync()
    {
        if (SelectedModListItem?.Metadata == null) return;
        var mod = SelectedModListItem.Metadata;
        
        // 1. 기본 종속성과 버전별 종속성을 모두 담을 집합(중복 방지)
        var allDependencies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        
        if (mod.About?.modDependencies != null)
        {
            foreach (var pid in mod.About.modDependencies)
                allDependencies.Add(pid);
        }

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

        if (requiredMods.Count == 0)
        {
            if (RequestShowAlert != null)
                await RequestShowAlert.Invoke("종속성이 선언되지 않았거나 설치된 모드 중 일치하는 종속성을 찾을 수 없습니다!");
            return;
        }
        
        // 실제 참조 모드 로드
        foreach (var refMod in ModLister.FindAllReferenceMods(mod))
        {
            if (!ReferenceMods.Contains(refMod))
                ReferenceMods.Add(refMod);
        }
        
        RefreshModList();
    }

    [RelayCommand]
    private async Task OpenExplorerAsync()
    {
        if (SelectedModListItem?.Metadata?.Path != null)
        {
            var path = SelectedModListItem.Metadata.Path;
            var targetMod = SelectedModListItem.Metadata;
            
            // 폴더 존재 여부를 뷰모델 단계에서 먼저 확인
            if (!System.IO.Directory.Exists(path))
            {
                if (RequestShowAlert != null)
                {
                    await RequestShowAlert.Invoke("존재하지 않는 폴더입니다.");
                }
                
                // 1. 뷰모델의 캐시 리스트에서 해당 모드 제거
                _allModsCached.Remove(targetMod);
                _officialModsCached.Remove(targetMod);
                _localModsCached.Remove(targetMod);
                _workshopModsCached.Remove(targetMod);
                
                // 2. 참조 모드로 선택되어 있었다면 제거
                if (ReferenceMods.Contains(targetMod))
                {
                    ReferenceMods.Remove(targetMod);
                }

                // 3. 현재 선택된 모드 초기화
                SelectedModListItem = null;
                SelectedModInfoText = "선택된 모드가 없습니다.";
                CanComplete = false;

                // 4. Core 엔진 레지스트리(RegisteredMods)에서 영구 삭제 및 자원 해제
                targetMod.Dispose();
                
                // 5. 남은 캐시를 바탕으로 UI 리스트 갱신
                RefreshModList(); 
                return;
            }

            _processService.OpenFolderInExplorer(path);
        }
    }

    private bool CanExecuteComplete() => CanComplete && SelectedMod != null;
    [RelayCommand(CanExecute = nameof(CanExecuteComplete))]
    private void Complete()
    {
        RequestClose?.Invoke();
    }
}