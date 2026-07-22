using CommunityToolkit.Mvvm.ComponentModel;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorGUI.ViewModels;

public partial class TranslationAnalyzerItemViewModel : ViewModelBase
{
    public TranslationAnalyzerEntry Entry { get; }

    [ObservableProperty] private bool _isChecked;
    [ObservableProperty] private string _modIdentifier = "UNKNOWN";
    [ObservableProperty] private string _filePath = string.Empty;
    [ObservableProperty] private int _originalCount;
    [ObservableProperty] private string _changesString = string.Empty;
    [ObservableProperty] private string _extractionMethod = "자동";
    [ObservableProperty] private string _saveMethodText = "덧붙이기";

    public TranslationAnalyzerItemViewModel(TranslationAnalyzerEntry entry)
    {
        Entry = entry;
        IsChecked = entry.HasChanges;
        ModIdentifier = entry.Metadata?.Identifier ?? "UNKNOWN";
        FilePath = entry.FilePath;
        OriginalCount = entry.OriginalTranslations.Count;
        ChangesString = entry.ChangesString;
        ExtractionMethod = entry.Metadata == null ? "지정 필요" : "자동";
        UpdateSaveMethodText();
    }

    public void UpdateFromEntry()
    {
        ModIdentifier = Entry.Metadata?.Identifier ?? "UNKNOWN";
        OriginalCount = Entry.OriginalTranslations.Count;
        ChangesString = Entry.ChangesString;
        ExtractionMethod = "수동";
        UpdateSaveMethodText();
    }

    public void UpdateSaveMethodText()
    {
        SaveMethodText = Entry.SaveMethod switch
        {
            TranslationAnalyzerEntry.SaveMethodEnum.Append => "덧붙이기",
            TranslationAnalyzerEntry.SaveMethodEnum.Overwrite => "재구성(덮어씌우기)",
            TranslationAnalyzerEntry.SaveMethodEnum.RewriteNewFile => "재구성(새파일)",
            TranslationAnalyzerEntry.SaveMethodEnum.New => "신규만 새파일",
            _ => "덧붙이기"
        };
    }
}