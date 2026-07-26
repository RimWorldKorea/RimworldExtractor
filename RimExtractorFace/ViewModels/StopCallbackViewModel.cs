using System;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace RimExtractorFace.ViewModels;

public partial class StopCallbackViewModel : ViewModelBase
{
    [ObservableProperty] private string _filePath = string.Empty;

    public bool OverwriteConfirmed { get; private set; } = false;
    public event Action? RequestClose;

    public StopCallbackViewModel(string path = "")
    {
        FilePath = path;
    }

    [RelayCommand]
    private void Overwrite()
    {
        OverwriteConfirmed = true;
        RequestClose?.Invoke();
    }

    [RelayCommand]
    private void Skip()
    {
        OverwriteConfirmed = false;
        RequestClose?.Invoke();
    }
}