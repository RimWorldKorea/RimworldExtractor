using CommunityToolkit.Mvvm.ComponentModel;

namespace RimExtractorFace.ViewModels;

public partial class ProcedureItemViewModel : ViewModelBase
{
    [ObservableProperty] private string _name;
    [ObservableProperty] private bool _isEnabled;

    public ProcedureItemViewModel(string name, bool isEnabled)
    {
        Name = name;
        IsEnabled = isEnabled;
    }
}