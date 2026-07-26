using Avalonia.Controls;
using MsBox.Avalonia;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

public partial class SettingsWindow : Window
{
    public SettingsWindow()
    {
        InitializeComponent();

        // 서비스 주입 제거, 뷰모델 단독 인스턴스화
        var viewModel = new SettingsViewModel();
        DataContext = viewModel;

        viewModel.RequestShowAlert += async (title, message) =>
        {
            var box = MessageBoxManager.GetMessageBoxStandard(title, message);
            await box.ShowAsync();
        };

        viewModel.RequestClose += () => Close();
    }
}