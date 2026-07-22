using Avalonia.Controls;
using Avalonia.Platform.Storage;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using RimworldExtractorGUI.Utils;
using RimworldExtractorGUI.ViewModels;
using RimworldExtractorInternal;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace RimworldExtractorGUI.Views;

public partial class MainWindow : Window
{
    private readonly MainWindowViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();

        _viewModel = new MainWindowViewModel();
        DataContext = _viewModel;

        // 로그 작성기 연결
        Log.Out = new AvaloniaLogWriter(TextBoxLog);

        // ViewModel 이벤트 핸들러 구독
        _viewModel.RequestSelectMod += OnSelectModRequestedAsync;
        _viewModel.RequestExtract += OnExtractRequestedAsync;
        _viewModel.RequestConvertXlsx += OnConvertXlsxRequestedAsync;
        _viewModel.RequestConvertXml += OnConvertXmlRequestedAsync;
        _viewModel.RequestOpenSettings += OnOpenSettingsRequestedAsync;
        _viewModel.RequestOpenJpgPackager += OnOpenJpgPackagerRequestedAsync;
        _viewModel.RequestOpenTranslationAnalyzer += OnOpenTranslationAnalyzerRequestedAsync;

        try
        {
            Prefabs.Load();
        }
        catch (Exception e)
        {
            ShowErrorMessageAndClose($"Prefabs.dat 파일의 버전이 구버전이거나 손상되었습니다. 파일 삭제 후 다시 진행해주세요.\n에러메시지: {e.Message}");
        }
    }

    private async void ShowErrorMessageAndClose(string message)
    {
        var box = MessageBoxManager.GetMessageBoxStandard("에러", message);
        await box.ShowAsync();
        Close();
    }

    private async Task OnSelectModRequestedAsync()
    {
        var win = new SelectModWindow();
        await win.ShowDialog(this);

        if (win.IsSuccess && win.SelectedMod != null)
        {
            _viewModel.UpdateSelectedModInfo(win.SelectedMod, win.SelectedFolders, win.ReferenceMods);
        }
    }

    private async Task OnExtractRequestedAsync()
    {
        if (_viewModel.SelectedMod == null || _viewModel.SelectedFolders == null || _viewModel.ReferenceMods == null)
            return;

        Log.Msg("추출 시작...");

        var extraction = Extractor.ExtractTranslationData(_viewModel.SelectedMod, _viewModel.SelectedFolders, _viewModel.ReferenceMods);
        var outPath = _viewModel.SelectedMod.Identifier.StripInvaildChars();

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

        string buildYamlText = RimworldExtractorInternal.Utils.WriteBuildYamlText(_viewModel.SelectedMod);
        File.WriteAllText(Path.Combine(outPath, "LoadFolders.Build.yaml"), buildYamlText);

        var (cntDefs, cntKeyed, cntStrings, cntPatches) = extraction.Count();
        Log.Msg($"번역 데이터 수: 총 {extraction.Count}개 중 Defs {cntDefs}개, Keyed {cntKeyed}개, Strings {cntStrings}개, Patches {cntPatches}개, 완료!");

        var box = MessageBoxManager.GetMessageBoxStandard("완료", "완료되었습니다! 추출된 파일의 위치를 탐색기로 열까요?", ButtonEnum.YesNo);
        if (await box.ShowAsync() == ButtonResult.Yes)
        {
            Process.Start(new ProcessStartInfo { FileName = outPath, UseShellExecute = true });
        }
    }

    private async Task OnConvertXlsxRequestedAsync()
    {
        var win = new XmlisterWindow();
        await win.ShowDialog(this);

        if (win.IsSuccess && win.FileNames.Length > 0)
        {
            var roots = win.FileNames;
            for (var i = 0; i < roots.Length; i++)
            {
                var root = roots[i];
                var translations = IO.FromLanguageXml(root);
                IO.ToExcel(translations, Path.Combine(root, Path.GetFileNameWithoutExtension(root)));
                Log.Msg($"{i + 1}/{roots.Length}::수정 완료: {root}");
            }

            var box = MessageBoxManager.GetMessageBoxStandard("완료", "변환이 완료되었습니다!");
            await box.ShowAsync();
        }
    }

    private async Task OnConvertXmlRequestedAsync()
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "림 추출기에서 생성한 엑셀 파일을 선택해주세요.",
            AllowMultiple = false,
            FileTypeFilter = new[] { new FilePickerFileType("번역 데이터 파일") { Patterns = new[] { "*.xlsx" } } }
        });

        if (files.Count > 0)
        {
            var path = files[0].Path.LocalPath;
            var translations = IO.FromExcel(path);
            IO.ToLanguageXml(translations, true, Prefabs.CommentOriginal, Path.GetFileName(path), Path.GetDirectoryName(path) ?? "");

            var box = MessageBoxManager.GetMessageBoxStandard("완료", "완료되었습니다! 변환된 폴더의 위치를 탐색기로 열까요?", ButtonEnum.YesNo);
            if (await box.ShowAsync() == ButtonResult.Yes)
            {
                Process.Start(new ProcessStartInfo { FileName = Path.GetDirectoryName(path) ?? "", UseShellExecute = true });
            }
        }
    }

    private async Task OnOpenSettingsRequestedAsync()
    {
        var settingsWin = new SettingsWindow();
        await settingsWin.ShowDialog(this);
    }

    private async Task OnOpenJpgPackagerRequestedAsync()
    {
        var win = new ImageFileCombinerWindow();
        await win.ShowDialog(this);
    }

    private async Task OnOpenTranslationAnalyzerRequestedAsync()
    {
        var pathSelectWin = new TranslationAnalyzerPathSelectWindow();
        await pathSelectWin.ShowDialog(this);

        if (pathSelectWin.IsSuccess && pathSelectWin.Paths.Length > 0)
        {
            var analyzerWin = new TranslationAnalyzerWindow(pathSelectWin.Paths);
            await analyzerWin.ShowDialog(this);
        }
    }
}