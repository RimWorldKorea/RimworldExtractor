using Avalonia.Controls;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

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