using Avalonia.Controls;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using RimworldExtractorGUI.Views;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorGUI.Services;

public interface IDialogService
{
    Task ShowAlertAsync(string title, string message);
    Task<bool> ConfirmAsync(string title, string message);
    Task<(bool IsSuccess, ModMetadata? SelectedMod, List<ExtractableFolder> SelectedFolders, List<ModMetadata> ReferenceMods)> ShowSelectModDialogAsync(ModMetadata? initialMod = null);
    Task<string[]?> ShowXmlisterDialogAsync();
    Task ShowSettingsDialogAsync();
    Task ShowImageFileCombinerDialogAsync();
    Task ShowTranslationAnalyzerDialogAsync();
}

public class DialogService : IDialogService
{
    private readonly Window _owner;

    public DialogService(Window owner)
    {
        _owner = owner;
    }

    public async Task ShowAlertAsync(string title, string message)
    {
        var box = MessageBoxManager.GetMessageBoxStandard(title, message);
        await box.ShowAsync();
    }

    public async Task<bool> ConfirmAsync(string title, string message)
    {
        var box = MessageBoxManager.GetMessageBoxStandard(title, message, ButtonEnum.YesNo);
        var result = await box.ShowAsync();
        return result == ButtonResult.Yes;
    }

    public async Task<(bool IsSuccess, ModMetadata? SelectedMod, List<ExtractableFolder> SelectedFolders, List<ModMetadata> ReferenceMods)> ShowSelectModDialogAsync(ModMetadata? initialMod = null)
    {
        var win = new SelectModWindow(initialMod);
        await win.ShowDialog(_owner);
        return (win.IsSuccess, win.SelectedMod, win.SelectedFolders, win.ReferenceMods);
    }

    public async Task<string[]?> ShowXmlisterDialogAsync()
    {
        var win = new XmlisterWindow();
        await win.ShowDialog(_owner);
        return win.IsSuccess ? win.FileNames : null;
    }

    public async Task ShowSettingsDialogAsync()
    {
        var win = new SettingsWindow();
        await win.ShowDialog(_owner);
    }

    public async Task ShowImageFileCombinerDialogAsync()
    {
        var win = new ImageFileCombinerWindow();
        await win.ShowDialog(_owner);
    }

    public async Task ShowTranslationAnalyzerDialogAsync()
    {
        var pathSelectWin = new TranslationAnalyzerPathSelectWindow();
        await pathSelectWin.ShowDialog(_owner);
        if (pathSelectWin.IsSuccess && pathSelectWin.Paths.Length > 0)
        {
            var analyzerWin = new TranslationAnalyzerWindow(pathSelectWin.Paths);
            await analyzerWin.ShowDialog(_owner);
        }
    }
}