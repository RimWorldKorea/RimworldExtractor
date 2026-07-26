using Avalonia.Controls;
using MsBox.Avalonia;
using RimExtractorCore.DataTypes;
using System.Collections.Generic;
using RimExtractorCore;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

public partial class TranslationAnalyzerWindow : Window
{
    private readonly TranslationAnalyzerViewModel _viewModel;

    public bool IsSuccess => _viewModel.IsSuccess;
    public IEnumerable<TranslationAnalyzerEntry> Entries => _viewModel.GetSelectedEntries();

    public TranslationAnalyzerWindow()
    {
        InitializeComponent();
        _viewModel = new TranslationAnalyzerViewModel();
        DataContext = _viewModel;
        InitEvents();
    }

    public TranslationAnalyzerWindow(string[] paths)
    {
        InitializeComponent();
        _viewModel = new TranslationAnalyzerViewModel(paths);
        DataContext = _viewModel;
        InitEvents();
    }

    private void InitEvents()
    {
        _viewModel.RequestShowAlert += async (title, msg) =>
        {
            var box = MessageBoxManager.GetMessageBoxStandard(title, msg);
            await box.ShowAsync();
        };

        _viewModel.RequestSelectModDialog += async (initMod) =>
        {
            var win = new SelectModWindow(initMod);
            await win.ShowDialog(this);
            return (win.IsSuccess, win.SelectedMod, win.SelectedFolders, win.ReferenceMods);
        };

        _viewModel.RequestClose += () => Close();
    }
}