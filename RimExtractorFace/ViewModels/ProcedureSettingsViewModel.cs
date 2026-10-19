using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.Input;
using RimExtractorCore;
using RimExtractorCore.Procedures;

namespace RimExtractorFace.ViewModels;

public partial class ProcedureSettingsViewModel : ViewModelBase
{
    public ObservableCollection<ProcedureItemViewModel> DefTreeProcedures { get; } = new();
    public ObservableCollection<ProcedureItemViewModel> ExtractionProcedures { get; } = new();
    public ObservableCollection<ProcedureItemViewModel> TranslationEntryProcedures { get; } = new();

    public event Action? RequestClose;

    public ProcedureSettingsViewModel()
    {
        LoadProcedures();
    }

    private void LoadProcedures()
    {
        // 1. DefTree 프로시저 로드
        var defDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "DefTree");
        var defProcs = ProcedureLoader.LoadProceduresFromDirectory<IXDocumentProcedure>(defDir);
        PopulateList(DefTreeProcedures, defProcs, SettingManager.Current.EnabledDefTreeProcedures);

        // 2. TranslationEntry 프로시저 로드
        var transDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "Translations");
        var transProcs = ProcedureLoader.LoadProceduresFromDirectory<ITranslationEntryProcedure>(transDir);
        PopulateList(TranslationEntryProcedures, transProcs, SettingManager.Current.EnabledTranslationEntryProcedures);

        // 3. Extraction 프로시저 로드
        var extDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Procedures", "Extractor");
        var extProcs = ProcedureLoader.LoadProceduresFromDirectory<IExtractionProcedure>(extDir);
        PopulateList(ExtractionProcedures, extProcs, SettingManager.Current.EnabledProcedures);

        // Extraction 단일 선택(라디오 버튼 같은) 로직 적용
        foreach (var item in ExtractionProcedures)
        {
            item.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(ProcedureItemViewModel.IsEnabled) && s is ProcedureItemViewModel { IsEnabled: true } currentItem)
                {
                    foreach (var other in ExtractionProcedures)
                    {
                        if (other != currentItem) other.IsEnabled = false;
                    }
                }
            };
        }
    }

    // 설정 파일에 기록된 순서를 우선하여 목록을 구성합니다.
    private void PopulateList<T>(ObservableCollection<ProcedureItemViewModel> targetList, List<T> loadedProcs, List<string> savedEnabledNames) where T : IProcedure
    {
        var remaining = loadedProcs.ToList();

        // 저장된 순서대로 먼저 배치하고 켜기
        foreach (var savedName in savedEnabledNames)
        {
            var match = remaining.FirstOrDefault(p => p.Name == savedName);
            if (match != null)
            {
                targetList.Add(new ProcedureItemViewModel(match.Name, true));
                remaining.Remove(match);
            }
        }

        // 설정에 없던 새로 감지된 프로시저는 맨 뒤에 비활성화 상태로 추가
        foreach (var proc in remaining)
        {
            targetList.Add(new ProcedureItemViewModel(proc.Name, false));
        }
    }
    
    // 1. 적용: 저장 로직만 실행하고 창은 닫지 않음
    [RelayCommand]
    private void Apply()
    {
        ApplySettings(SettingManager.Current);
    }

    // 2. 저장: 현재는 새 UI 기능 추가 없이 분리만 하므로 '적용'과 동일한 저장 로직 실행
    [RelayCommand]
    private void Save()
    {
        ApplySettings(SettingManager.Current);
        ApplySettings(SettingManager.LastSaved);
        SettingManager.SaveToDisk();
    }

    // 3. 닫기: 저장 없이 창만 닫음
    [RelayCommand]
    private void Close()
    {
        RequestClose?.Invoke();
    }
    
    private void ApplySettings(Settings target)
    {
        target.EnabledDefTreeProcedures = DefTreeProcedures.Where(p 
            => p.IsEnabled).Select(p => p.Name).ToList();
        
        target.EnabledTranslationEntryProcedures = TranslationEntryProcedures.Where(p 
            => p.IsEnabled).Select(p => p.Name).ToList();
        
        target.EnabledProcedures = ExtractionProcedures.Where(p 
            => p.IsEnabled).Select(p => p.Name).ToList();
    }
}