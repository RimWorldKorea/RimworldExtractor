using Avalonia.Controls;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

public partial class XmlToSheetWindow : Window
{
    private readonly XmlToSheetViewModel _viewModel;

    public string[] FileNames => _viewModel.FileNames;
    public bool IsSuccess => _viewModel.IsSuccess;

    public XmlToSheetWindow()
    {
        InitializeComponent();

        _viewModel = new XmlToSheetViewModel();
        DataContext = _viewModel;

        _viewModel.RequestClose += () => Close();
    }
}