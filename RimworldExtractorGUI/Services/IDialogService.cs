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