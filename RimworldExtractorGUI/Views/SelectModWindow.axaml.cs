using Avalonia.Controls;
using Avalonia.Input;

using MsBox.Avalonia;
using RimworldExtractorGUI.ViewModels;
using RimworldExtractorInternal.DataTypes;
using System.Collections.Generic;

namespace RimworldExtractorGUI.Views;

public partial class SelectModWindow : Window
{
    private readonly SelectModViewModel _viewModel;

    public ModMetadata? SelectedMod => _viewModel.SelectedMod;
    public List<ExtractableFolder> SelectedFolders { get; } = new();
    public List<ModMetadata> ReferenceMods => _viewModel.ReferenceMods;
    public bool IsSuccess { get; private set; } = false;

    public SelectModWindow() : this(null) { }
    
    public SelectModWindow(ModMetadata? initialMod = null)
    {
        InitializeComponent();

        _viewModel = new SelectModViewModel(initialMod);
        DataContext = _viewModel;

        _viewModel.RequestShowAlert += async (msg) =>
        {
            var box = MessageBoxManager.GetMessageBoxStandard("알림", msg);
            await box.ShowAsync();
        };

        _viewModel.RequestClose += () =>
        {
            SelectedFolders.Clear();
            if (ListBoxFolders.SelectedItems != null)
            {
                foreach (var item in ListBoxFolders.SelectedItems)
                {
                    if (item is ExtractableFolder folder)
                        SelectedFolders.Add(folder);
                }
            }

            IsSuccess = true;
            Close();
        };
    }

    private void Window_KeyDown(object? sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.A:
                _viewModel.OpenExplorerCommand.Execute(null);
                break;
            case Key.S:
                _viewModel.ToggleRefModCommand.Execute(null);
                break;
            case Key.D:
                _viewModel.IsFilterSelectedOnly = !_viewModel.IsFilterSelectedOnly;
                break;
        }
    }
}