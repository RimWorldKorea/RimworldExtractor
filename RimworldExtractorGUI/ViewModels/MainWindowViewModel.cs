using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimworldExtractorGUI.Services;
using RimworldExtractorInternal.Core;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorGUI.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly IDialogService _dialogService;
    private readonly IStorageService _storageService;
    private readonly IUpdateCheckService _versionService;
    private readonly IExtractionService _extractionService;
    private readonly IExternalProcessService _processService;

    [ObservableProperty]
    private string _selectedModsText = "모드를 선택해주세요.\n\n=== 사용 순서 ===\n1) '1. 모드 선택' 버튼 클릭\n2) '2. 번역 데이터 추출' 버튼 클릭\n3) 언어팩 생성 완료!\n5) 즐거운 림월드 되세요!\n\n제보 및 문의: 디스코드";

    [ObservableProperty]
    private string _versionText = "버전 확인 중...";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ExtractCommand))]
    private bool _canExtract = false;

    public ModMetadata? SelectedMod { get; private set; }
    public List<ExtractableFolder>? SelectedFolders { get; private set; }
    public List<ModMetadata>? ReferenceMods { get; private set; }

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
            // 비동기 버전 확인
            var latest = await _versionService.GetLatestVersionAsync();
            var current = _versionService.CurrentVersion;
            VersionText = latest == current
                ? $"{current} (최신 버전)"
                : $"{current} < {latest} (업데이트 필요)";
        }
        catch (Exception e)
        {
            Log.Wrn($"버전 확인 중 에러: {e.Message}");
            VersionText = "버전 확인 실패";
        }
    }

    public void UpdateSelectedModInfo(ModMetadata mod, List<ExtractableFolder> folders, List<ModMetadata> refMods)
    {
        SelectedMod = mod;
        SelectedFolders = folders;
        ReferenceMods = refMods.Except(new[] { mod }).ToList();
        CanExtract = true;

        var text = $"선택된 모드 : {SelectedMod.ModName}";
        if (ReferenceMods?.Count > 0)
        {
            var concatText = string.Join(", ", ReferenceMods.Select(x => x.ModName));
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
            UpdateSelectedModInfo(result.SelectedMod, result.SelectedFolders, result.ReferenceMods);
        }
    }

    [RelayCommand(CanExecute = nameof(CanExtract))]
    private async Task ExtractAsync()
    {
        if (SelectedMod == null || SelectedFolders == null || ReferenceMods == null) return;

        Log.Msg("추출 시작...");

        // UI 멈춤 방지를 위해 Task.Run 내부적으로 I/O 및 추출을 수행하는 서비스 호출
        var summary = await _extractionService.ExtractAndSaveAsync(SelectedMod, SelectedFolders, ReferenceMods);

        // 콘솔 출력 - (Defs, Keyed, Strings, Patches)
        Log.Msg($"추출 완료! 총 {summary.TotalCount}개 노드 (Defs {summary.DefsCount}개, Keyed {summary.KeyedCount}개, Strings {summary.StringsCount}개, Patches {summary.PatchesCount}개) 추출됨!");

        if (await _dialogService.ConfirmAsync("추출 완료", "결과 폴더를 열어보시겠습니까?"))
        {
            _processService.OpenFolderInExplorer(summary.OutputPath);
        }
    }

    [RelayCommand]
    private async Task ConvertXlsxAsync()
    {
        var fileNames = await _dialogService.ShowXmlisterDialogAsync();
        if (fileNames != null && fileNames.Length > 0)
        {
            await _extractionService.ConvertXmlToXlsxAsync(fileNames);
            await _dialogService.ShowAlertAsync("작업 완료", "변환이 완료되었습니다!");
        }
    }

    [RelayCommand]
    private async Task ConvertXmlAsync()
    {
        var path = await _storageService.OpenFileAsync("변환할 Excel 파일 선택", "Excel 파일", "*.xlsx");
        if (!string.IsNullOrEmpty(path))
        {
            await _extractionService.ConvertXlsxToXmlAsync(path);
            if (await _dialogService.ConfirmAsync("변환 완료", "결과 폴더를 열어보시겠습니까?"))
            {
                _processService.OpenFolderInExplorer(Path.GetDirectoryName(path) ?? "");
            }
        }
    }

    [RelayCommand]
    private async Task OpenSettingsAsync() => await _dialogService.ShowSettingsDialogAsync();

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