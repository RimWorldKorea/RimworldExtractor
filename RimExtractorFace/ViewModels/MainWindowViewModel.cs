using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimExtractorCore;
using RimExtractorCore.DataTypes;
using RimExtractorFace.Services;


namespace RimExtractorFace.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly IDialogService _dialogService;
    private readonly IStorageService _storageService;
    private readonly IUpdateCheckService _versionService;
    private readonly IExtractionService _extractionService;
    private readonly IExternalProcessService _processService;

    [ObservableProperty]
    private string _selectedModsText = "▸ 모드를 선택하면 추출 버튼이 활성화됩니다.\n\n▸ 번역 후 기본 검증 기능을 사용하여 결과물을 점검해보세요.\n\n▸ 안내: 현재 번역 분석기는 동작하지 않습니다.";

    [ObservableProperty]
    private string _versionText = "버전 확인 중...";

    [ObservableProperty]
    private bool? _isLatest;
    
    [ObservableProperty]
    private string _isSpecialBuildIfNotEmpty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ExtractCommand))]
    private bool _canExtract = false;

    public Mod? SelectedMod { get; private set; }
    public List<Mod>? ReferenceMods { get; private set; }

    // 생성자를 통해 의존성(Service) 주입
    public MainWindowViewModel(
        IDialogService dialogService,
        IStorageService storageService,
        IUpdateCheckService versionService,
        IExtractionService extractionService,
        IExternalProcessService processService)
    {
        _dialogService = dialogService;
        _storageService = storageService;
        _versionService = versionService;
        _extractionService = extractionService;
        _processService = processService;

        CheckVersionAsync();
    }

    private async void CheckVersionAsync()
    {
        try
        {
            var latestRaw = await _versionService.GetLatestVersionAsync();
            var latest = Version.Parse(latestRaw);
            
            var currentRaw = _versionService.CurrentVersion.Split("+").First().Split("-");
            var current = Version.Parse(currentRaw.First());
            if (currentRaw.Length > 1) IsSpecialBuildIfNotEmpty = currentRaw.Last();
            
            if (current >= latest)
            {
                IsLatest = true;
                VersionText = $"{current} 최신";
            }
            
            if (current < latest)
            {
                IsLatest = false;
                VersionText = $"{current} < {latestRaw} 업데이트 필요";
            }
        }
        catch (Exception e)
        {
            Log.Wrn($"버전 확인 중 에러\n\t{e.Message}");
            IsLatest = null;
            VersionText = "버전 확인 실패";
        }
    }

    public void UpdateSelectedModInfo(Mod mod, List<Mod> refMods)
    {
        SelectedMod = mod;
        ReferenceMods = refMods.Except(new[] { mod }).ToList();
        CanExtract = true;

        var text = $"선택된 모드 : {SelectedMod.Name}";
        if (ReferenceMods?.Count > 0)
        {
            var concatText = string.Join(", ", ReferenceMods.Select(x => x.Name));
            var stripedText = concatText.Substring(0, Math.Min(concatText.Length, 200));
            if (concatText.Length > 200) stripedText += "...";
            text += $"\n기준 모드 : {stripedText}";
        }
        SelectedModsText = text;
    }

    [RelayCommand]
    private async Task SelectModAsync()
    {
        var result = await _dialogService.ShowSelectModDialogAsync(SelectedMod);
        if (result.IsSuccess && result.SelectedMod != null)
        {
            UpdateSelectedModInfo(result.SelectedMod, result.ReferenceMods);
        }
    }

    [RelayCommand(CanExecute = nameof(CanExtract))]
    private async Task ExtractAsync()
    {
        if (SelectedMod == null || ReferenceMods == null) return;

        Log.Msg("추출 시작...");

        try
        {
            // UI 멈춤 방지를 위해 Task.Run 내부적으로 I/O 및 추출을 수행하는 서비스 호출
            var summary = await _extractionService.ExtractAndSaveAsync(SelectedMod, ReferenceMods);
            
            // 콘솔 출력 - (Defs, Keyed, Strings, Patches)
            Log.Good($"추출 완료.\n\tDefs: {summary.DefsCount}\n\tKeyed: {summary.KeyedCount}\n\tStrings: {summary.StringsCount}\n\tPatches: {summary.PatchesCount}\n\tTotal: {summary.TotalCount}", true);

            if (await _dialogService.ConfirmAsync("추출 완료", "저장한 폴더를 열어보시겠습니까?"))
            {
                _processService.OpenFolderInExplorer(summary.OutputPath);
            }
        }
        catch (Exception e)
        {
            // 🚨 코어 엔진에서 던진 예외를 여기서 잡아내어 프로그램 크래시 방지!
            Log.Err(e.Message);
            await _dialogService.ShowAlertAsync("추출 실패", e.Message);
        }

    }

    [RelayCommand]
    private async Task ConvertToSheetAsync()
    {
        var fileNames = await _dialogService.ShowXmlListerDialogAsync();
        if (fileNames != null && fileNames.Length > 0)
        {
            await _extractionService.ConvertXmlToOdsAsync(fileNames);
            await _dialogService.ShowAlertAsync("작업 완료", "변환이 완료되었습니다!");
        }
    }

    [RelayCommand]
    private async Task ConvertToXmlAsync()
    {
        var path = await _storageService.OpenFileAsync("변환할 스프레드시트 선택", "스프레드시트", "*.xlsx", "*.ods");
        if (!string.IsNullOrEmpty(path))
        {
            await _extractionService.ConvertSheetToXmlAsync(path);
            if (await _dialogService.ConfirmAsync("변환 완료", "결과 폴더를 열어보시겠습니까?"))
            {
                _processService.OpenFolderInExplorer(Path.GetDirectoryName(path) ?? "");
            }
        }
    }
    
    [RelayCommand]
    private async Task InspectTranslationAsync()
    {
        // 폴더 선택 창 호출
        var path = await _storageService.OpenFolderAsync("검증할 번역의 Languages 폴더가 위치한 폴더를 지정해주세요.");
        if (string.IsNullOrEmpty(path)) return;

        string translationLanguageString = SettingManager.Current.TranslationLanguage;

        if (string.IsNullOrEmpty(translationLanguageString))
        {
            Log.Err("목표 언어가 설정되지 않음.");
            return;
        }
        
        if (!LanguageInfo.TryParse(translationLanguageString, out LanguageInfo? translationLanguage))
        {
            Log.Err($"\"{translationLanguageString}\"(은)는 유효한 언어가 아닙니다.");
            return;
        }

        await Task.Run(() =>
            RimExtractorCore.TranslationInspector.TranslationInspector.CheckTranslation(path, translationLanguage!)
        );

    }

    [RelayCommand]
    private async Task OpenSettingsAsync() => await _dialogService.ShowSettingsDialogAsync();

    [RelayCommand]
    private async Task OpenProcedureSettingsAsync() => await _dialogService.ShowProcedureSettingsDialogAsync();
    
    [RelayCommand]
    private async Task OpenJpgPackagerAsync() => await _dialogService.ShowImageFileCombinerDialogAsync();

    [RelayCommand]
    private async Task OpenTranslationAnalyzerAsync() => await _dialogService.ShowTranslationAnalyzerDialogAsync();

    [RelayCommand]
    private void OpenVersionUrl()
    {
        // 깃허브 릴리즈 URL 열기
        _processService.OpenUrlInBrowser(_versionService.LatestUrl);
    }

    [RelayCommand]
    private void OpenDiscordUrl()
    {
        _processService.OpenUrlInBrowser(_versionService.DiscordUrl);
    }
}