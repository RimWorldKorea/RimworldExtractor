using Avalonia.Controls;
using MsBox.Avalonia;
using RimworldExtractorGUI.ViewModels;

namespace RimworldExtractorGUI.Views;

public partial class SettingsWindow : Window
{
    public SettingsWindow()
    {
        InitializeComponent();

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