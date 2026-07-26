using Avalonia.Controls;
using RimExtractorCore;
using MsBox.Avalonia;
using RimExtractorFace.Services;
using RimExtractorFace.Utils;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

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

        _viewModel = new MainWindowViewModel(
            dialogService,
            storageService,
            versionService,
            extractionService,
            processService);

        DataContext = _viewModel;

        // 2. 로그 라이터 연결
        var logPrinter = new LogPrinter(TextBoxLog);
        Log.Out = logPrinter;

        // 기존 큐에 쌓여있던 완성된 로그 문자열들을 직접 printer에 밀어넣기
        foreach (var cachedMsg in Log.Messages)
        {
            logPrinter.WriteLine(cachedMsg);
        }
    }

    private async void ShowErrorMessageAndClose(string message)
    {
        var box = MessageBoxManager.GetMessageBoxStandard("에러 발생", message);
        await box.ShowAsync();
        Close();
    }
}