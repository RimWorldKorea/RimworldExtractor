using System.Diagnostics;
using RimExtractorCore;

namespace RimExtractorFace.Services;

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
        
        try
        {
            Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
        }
        catch (Exception e)
        {
            Log.Wrn($"브라우저 실행 실패 ({url}): {e.Message}");
        }
    }

    public void OpenFolderInExplorer(string folderPath)
    {
        if (string.IsNullOrWhiteSpace(folderPath)) return;
        
        if (!Directory.Exists(folderPath))
        {
            Log.Wrn($"존재하지 않는 폴더입니다: {folderPath}");
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo { FileName = folderPath, UseShellExecute = true });
        }
        catch (Exception e)
        {
            Log.Err($"탐색기 실행 실패 ({folderPath}): {e.Message}");
        }
    }

    public void OpenFileWithDefaultApp(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return;

        // [추가] 파일이 실제로 존재하는지 1차 검증
        if (!File.Exists(filePath))
        {
            Log.Wrn($"존재하지 않는 파일입니다: {filePath}");
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
        }
        catch (Exception e)
        {
            Log.Err($"파일 실행 실패 ({filePath}): {e.Message}");
        }
    }
}