using Avalonia.Controls;
using RimworldExtractorGUI.Services;
using RimworldExtractorGUI.Utils;
using RimworldExtractorGUI.ViewModels;
using RimworldExtractorInternal;
using MsBox.Avalonia;

namespace RimworldExtractorGUI.Views;

public partial class MainWindow : Window
{
    private readonly MainWindowViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();

        // 1. 서비스 초기화 및 주입
        var dialogService = new AvaloniaDialogService(this);
        var storageService = new AvaloniaStorageService(this);
        var versionService = new GitHubVersionCheckService();
        _viewModel = new MainWindowViewModel(dialogService, storageService, versionService);
        DataContext = _viewModel;

        // 2. 로그 라이터 설정
        Log.Out = new AvaloniaLogWriter(TextBoxLog);

        // 3. 설정 파일 로드
        try
        {
            Prefabs.Load();
        }
        catch (Exception e)
        {
            ShowErrorMessageAndClose($"Prefabs.dat 로드 실패 : {e.Message}");
        }
    }

    private async void ShowErrorMessageAndClose(string message)
    {
        var box = MessageBoxManager.GetMessageBoxStandard("오류", message);
        await box.ShowAsync();
        Close();
    }
}