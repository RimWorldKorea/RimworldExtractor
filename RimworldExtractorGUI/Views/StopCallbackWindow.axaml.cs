using Avalonia.Controls;
using RimworldExtractorGUI.ViewModels;

namespace RimworldExtractorGUI.Views;

public partial class StopCallbackWindow : Window
{
    private readonly StopCallbackViewModel _viewModel;

    public bool OverwriteConfirmed => _viewModel.OverwriteConfirmed;

    public StopCallbackWindow() : this("") { }
    
    public StopCallbackWindow(string path = "")
    {
        InitializeComponent();

        _viewModel = new StopCallbackViewModel(path);
        DataContext = _viewModel;

        _viewModel.RequestClose += () => Close();
    }
}