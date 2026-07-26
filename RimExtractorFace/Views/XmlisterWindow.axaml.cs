using Avalonia.Controls;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

public partial class XmlisterWindow : Window
{
    private readonly XmlisterViewModel _viewModel;

    public string[] FileNames => _viewModel.FileNames;
    public bool IsSuccess => _viewModel.IsSuccess;

    public XmlisterWindow()
    {
        InitializeComponent();

        _viewModel = new XmlisterViewModel();
        DataContext = _viewModel;

        _viewModel.RequestClose += () => Close();
    }
}