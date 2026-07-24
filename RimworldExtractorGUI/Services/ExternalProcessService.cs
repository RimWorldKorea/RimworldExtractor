using System.Diagnostics;

namespace RimworldExtractorGUI.Services;

public interface IExternalProcessService
{
    void OpenUrlInBrowser(string url);
    void OpenFolderInExplorer(string folderPath);
    void OpenFileWithDefaultApp(string filePath);
}

public class ExternalProcessService : IExternalProcessService
{
    public void OpenUrlInBrowser(string url)
    {
        if (string.IsNullOrWhiteSpace(url)) return;
        Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
    }

    public void OpenFolderInExplorer(string folderPath)
    {
        if (string.IsNullOrWhiteSpace(folderPath)) return;
        Process.Start(new ProcessStartInfo { FileName = folderPath, UseShellExecute = true });
    }

    public void OpenFileWithDefaultApp(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return;
        Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
    }
}