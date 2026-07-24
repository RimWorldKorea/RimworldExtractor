using Avalonia.Controls;
using MsBox.Avalonia;
using RimworldExtractorGUI.Services;
using RimworldExtractorGUI.ViewModels;

namespace RimworldExtractorGUI.Views;

public partial class SettingsWindow : Window
{
    public SettingsWindow()
    {
        InitializeComponent();

        // 설정 관리 서비스 주입
        var viewModel = new SettingsViewModel(new PrefabSettingsService());
        DataContext = viewModel;

        viewModel.RequestShowAlert += async (title, message) =>
        {
            var box = MessageBoxManager.GetMessageBoxStandard("알림", message);
            await box.ShowAsync();
        };

        viewModel.RequestClose += () => Close();
    }
}