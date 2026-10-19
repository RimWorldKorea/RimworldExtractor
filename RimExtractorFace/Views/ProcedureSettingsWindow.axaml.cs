using System.Collections.ObjectModel;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.VisualTree;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

public partial class ProcedureSettingsWindow : Window
{
    private static readonly DataFormat<ProcedureItemViewModel> ProcItemFormat =
        DataFormat.CreateInProcessFormat<ProcedureItemViewModel>("ProcItem");
    private static readonly DataFormat<ListBox> SourceBoxFormat =
        DataFormat.CreateInProcessFormat<ListBox>("SourceBox");

    public ProcedureSettingsWindow()
    {
        InitializeComponent();
        
        var viewModel = new ProcedureSettingsViewModel();
        DataContext = viewModel;

        viewModel.RequestClose += () => Close();

        AddHandler(DragDrop.DropEvent, OnDrop);
        AddHandler(DragDrop.DragOverEvent, OnDragOver);
    }

    private async void OnPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (sender is Control control && control.DataContext is ProcedureItemViewModel item)
        {
            var listBox = control.GetVisualAncestors().OfType<ListBox>().FirstOrDefault();
            if (listBox == null) return;
            
            var transferItem = new DataTransferItem();
            transferItem.Set(ProcItemFormat, item);
            transferItem.Set(SourceBoxFormat, listBox);

            var dragData = new DataTransfer();
            dragData.Add(transferItem);
            
            await DragDrop.DoDragDropAsync(e, dragData, DragDropEffects.Move);
        }
    }

    private void OnDragOver(object? sender, DragEventArgs e)
    {
        e.DragEffects = DragDropEffects.Move;
    }

    private void OnDrop(object? sender, DragEventArgs e)
    {
        var sourceItem = e.DataTransfer.TryGetValue(ProcItemFormat);
        var sourceBox = e.DataTransfer.TryGetValue(SourceBoxFormat);

        if (sourceItem != null && sourceBox != null)
        {
            var targetControl = e.Source as Control;
            var targetItem = targetControl?.DataContext as ProcedureItemViewModel;

            if (targetItem != null && targetItem != sourceItem)
            {
                if (sourceBox.ItemsSource is ObservableCollection<ProcedureItemViewModel> list)
                {
                    int sourceIndex = list.IndexOf(sourceItem);
                    int targetIndex = list.IndexOf(targetItem);

                    if (sourceIndex >= 0 && targetIndex >= 0)
                    {
                        list.Move(sourceIndex, targetIndex);
                    }
                }
            }
        }
    }
}