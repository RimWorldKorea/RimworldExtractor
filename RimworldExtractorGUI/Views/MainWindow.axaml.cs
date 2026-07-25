using Avalonia.Controls;
using RimworldExtractorGUI.Services;
using RimworldExtractorGUI.Utils;
using RimworldExtractorGUI.ViewModels;
using RimworldExtractorInternal.Core;
using MsBox.Avalonia;

namespace RimworldExtractorGUI.Views;

public partial class MainWindow : Window
{
    private readonly MainWindowViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();

        // 1. 서비스 초기화 및 주입
        var dialogService = new DialogService(this);
        var storageService = new StorageService(this);
        var versionService = new UpdateCheckService();
        var extractionService = new ExtractionService();
        var processService = new ExternalProcessService();
        var settingsService = new PrefabSettingsService();

        _viewModel = new MainWindowViewModel(
            dialogService,
            storageService,
            versionService,
            extractionService,
            processService);

        DataContext = _viewModel;

        // 2. 로그 라이터 연결
        Log.Out = new LogPrinter(TextBoxLog);

        // 3. 앱 구동 시 Prefabs 로드
        try
        {
            settingsService.Load();
        }
        catch (Exception e)
        {
            ShowErrorMessageAndClose($"Prefabs.dat을 읽는 중 에러 발생: {e.Message}");
        }
    }

    private async void ShowErrorMessageAndClose(string message)
    {
        var box = MessageBoxManager.GetMessageBoxStandard("에러 발생", message);
        await box.ShowAsync();
        Close();
    }
}