using Avalonia.Controls;
using RimworldExtractorGUI.ViewModels;

namespace RimworldExtractorGUI.Views;

public partial class TranslationAnalyzerPathSelectWindow : Window
{
    private readonly TranslationAnalyzerPathSelectViewModel _viewModel;

    public string[] Paths => _viewModel.Paths;
    public bool IsSuccess => _viewModel.IsSuccess;

    public TranslationAnalyzerPathSelectWindow()
    {
        InitializeComponent();

        _viewModel = new TranslationAnalyzerPathSelectViewModel();
        DataContext = _viewModel;

        _viewModel.RequestClose += () => Close();
    }
}