using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

public partial class InitialPathSelectWindow : Window
{
    public InitialPathSelectWindow()
    {
        InitializeComponent();

        var viewModel = new InitialPathSelectViewModel();
        DataContext = viewModel;

        // ViewModel에서 창 닫기 요청 시 처리
        viewModel.RequestClose += () =>
        {
            if (Avalonia.Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                var mainWindow = new MainWindow();
                desktop.MainWindow = mainWindow;
                mainWindow.Show();
                Close();
            }
        };
    }
}