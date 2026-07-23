using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimworldExtractorGUI.Services;
using RimworldExtractorGUI.Utils;
using RimworldExtractorInternal;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorGUI.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly IDialogService _dialogService;
    private readonly IStorageService _storageService;
    private readonly IVersionCheckService _versionService;

    [ObservableProperty]
    private string _selectedModsText = "모드가 선택되지 않았습니다.\n\n=== 사용 방법 ===\n1) '1. 모드 선택' 버튼을 누릅니다.\n2) '2. 추출' 버튼을 누릅니다.\n3) 변환이 완료되면 출력 디렉토리가 열립니다.\n5) 번역을 시작하세요!\n\n※ 모드를 선택하면 모드 정보가 이곳에 표시됩니다.";

    [ObservableProperty]
    private string _versionText = "버전 확인 중...";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ExtractCommand))]
    private bool _canExtract = false;

    public ModMetadata? SelectedMod { get; private set; }
    public List<ExtractableFolder>? SelectedFolders { get; private set; }
    public List<ModMetadata>? ReferenceMods { get; private set; }

    // 생성자를 통해 Service 주입
    public MainWindowViewModel(IDialogService dialogService, IStorageService storageService, IVersionCheckService versionService)
    {
        _dialogService = dialogService;
        _storageService = storageService;
        _versionService = versionService;
        
        CheckVersionAsync();
    }

    private async void CheckVersionAsync()
    {
        try
        {
            // 비동기로 안전하게 최신 버전 호출
            var latest = await _versionService.GetLatestVersionAsync();
            var current = _versionService.CurrentVersion;

            VersionText = latest == current
                ? $"{current} (최신 버전)"
                : $"{current} < {latest} (업데이트 가능)";
        }
        catch (Exception e)
        {
            Log.Wrn($"버전 확인 실패 : {e.Message}");
            VersionText = "버전 확인 실패";
        }
    }

    public void UpdateSelectedModInfo(ModMetadata mod, List<ExtractableFolder> folders, List<ModMetadata> refMods)
    {
        SelectedMod = mod;
        SelectedFolders = folders;
        ReferenceMods = refMods.Except(new[] { mod }).ToList();
        CanExtract = true;

        var text = $"대상 모드 : {SelectedMod.ModName}";
        if (ReferenceMods?.Count > 0)
        {
            var concatText = string.Join(", ", ReferenceMods.Select(x => x.ModName));
            var stripedText = concatText.Substring(0, Math.Min(concatText.Length, 200));
            if (concatText.Length > 200) stripedText += "...";
            text += $"\n참조 모드 : {stripedText}";
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
        
        // UI가 멈추지 않도록 무거운 추출 로직은 Task.Run으로 감쌉니다.
        var extraction = await Task.Run(() => 
            Extractor.ExtractTranslationData(SelectedMod, SelectedFolders, ReferenceMods));

        var outPath = SelectedMod.Identifier.StripInvaildChars();

        // I/O 작업 수행
        await Task.Run(() =>
        {
            switch (Prefabs.Method)
            {
                case Prefabs.ExtractionMethod.Excel:
                    IO.ToExcel(extraction, Path.Combine(outPath, outPath));
                    break;
                case Prefabs.ExtractionMethod.Languages:
                    IO.ToLanguageXml(extraction, false, false, outPath, outPath);
                    break;
                case Prefabs.ExtractionMethod.LanguagesWithComments:
                    IO.ToLanguageXml(extraction, false, true, outPath, outPath);
                    break;
            }

            string buildYamlText = RimworldExtractorInternal.Utils.WriteBuildYamlText(SelectedMod);
            File.WriteAllText(Path.Combine(outPath, "LoadFolders.Build.yaml"), buildYamlText);
        });

        // 튜플 요소 추출 - (Defs, Keyed, Strings, Patches)
        var (cntDefs, cntKeyed, cntStrings, cntPatches) = RimworldExtractorInternal.Utils.Count(extraction);
        Log.Msg($"번역 데이터 수: 총 {extraction.Count}개 중 Defs {cntDefs}개, Keyed {cntKeyed}개, Strings {cntStrings}개, Patches {cntPatches}개 완료!");

        if (await _dialogService.ConfirmAsync("추출 완료", "추출된 폴더를 열어보시겠습니까?"))
        {
            Process.Start(new ProcessStartInfo { FileName = outPath, UseShellExecute = true });
        }
    }

    [RelayCommand]
    private async Task ConvertXlsxAsync()
    {
        var fileNames = await _dialogService.ShowXmlisterDialogAsync();
        if (fileNames != null && fileNames.Length > 0)
        {
            await Task.Run(() =>
            {
                for (var i = 0; i < fileNames.Length; i++)
                {
                    var root = fileNames[i];
                    var translations = IO.FromLanguageXml(root);
                    IO.ToExcel(translations, Path.Combine(root, Path.GetFileNameWithoutExtension(root)));
                    Log.Msg($"{i + 1}/{fileNames.Length}::변환 완료 : {root}");
                }
            });

            await _dialogService.ShowAlertAsync("알림", "모든 파일의 변환이 완료되었습니다!");
        }
    }

    [RelayCommand]
    private async Task ConvertXmlAsync()
    {
        var path = await _storageService.OpenFileAsync("변환할 Excel 파일 선택", "Excel 파일", "*.xlsx");
        if (!string.IsNullOrEmpty(path))
        {
            await Task.Run(() =>
            {
                var translations = IO.FromExcel(path);
                IO.ToLanguageXml(translations, true, Prefabs.CommentOriginal, Path.GetFileName(path), Path.GetDirectoryName(path) ?? "");
            });

            if (await _dialogService.ConfirmAsync("변환 완료", "변환된 폴더를 열어보시겠습니까?"))
            {
                Process.Start(new ProcessStartInfo { FileName = Path.GetDirectoryName(path) ?? "", UseShellExecute = true });
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
        // 서비스에서 URL을 가져옴
        Process.Start(new ProcessStartInfo { FileName = _versionService.LatestUrl, UseShellExecute = true });
    }

    [RelayCommand]
    private void OpenDiscordUrl()
    {
        Process.Start(new ProcessStartInfo { FileName = _versionService.DiscordUrl, UseShellExecute = true });
    }
}