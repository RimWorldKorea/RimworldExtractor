using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimExtractorFace.Utils;

namespace RimExtractorFace.ViewModels;

public partial class ImageFileCombinerViewModel : ViewModelBase
{
    [ObservableProperty] private string _pathImage = string.Empty;
    [ObservableProperty] private string _pathFile = string.Empty;

    public event Func<string, string, Task<string?>>? RequestSaveFilePicker;
    public event Func<string, string, Task>? RequestShowAlert;
    public event Func<string, string, Task<bool>>? RequestConfirm;

    [RelayCommand]
    private async Task SelectPathImageAsync(IStorageProvider storageProvider)
    {
        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "이미지 파일을 지정해주세요.",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("이미지 파일") { Patterns = new[] { "*.jpg", "*.png", "*.gif" } }
            }
        });

        if (files.Count > 0)
        {
            PathImage = files[0].Path.LocalPath;
        }
    }

    [RelayCommand]
    private async Task SelectPathFileAsync(IStorageProvider storageProvider)
    {
        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "패키징할 압축파일의 경로를 지정해주세요.",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("ZIP 압축파일") { Patterns = new[] { "*.zip" } }
            }
        });

        if (files.Count > 0)
        {
            PathFile = files[0].Path.LocalPath;
        }
    }

    [RelayCommand]
    private async Task SelectPathDirAsync(IStorageProvider storageProvider)
    {
        var folders = await storageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "패키징할 폴더의 경로를 지정해주세요.",
            AllowMultiple = false
        });

        if (folders.Count > 0)
        {
            PathFile = folders[0].Path.LocalPath;
        }
    }

    [RelayCommand]
    private async Task ProcessPackageAsync()
    {
        var filePath = PathFile ?? string.Empty;
        var imgPath = string.IsNullOrEmpty(PathImage) ? null : PathImage;
        var imgExtension = Path.GetExtension(imgPath) ?? ".jpg";

        if (imgPath != null && !File.Exists(imgPath))
        {
            if (RequestShowAlert != null)
            {
                await RequestShowAlert.Invoke("에러",
                    "경로 상에 이미지 파일이 존재하지 않거나 엑세스 권한이 없습니다.\n" +
                    "파일이 존재함에도 에러가 발생한다면 관리자 권한으로 실행하거나, 파일을 다른 위치로 옮긴 후 다시 시도해주세요.");
            }
            return;
        }

        if (File.Exists(filePath))
        {
            if (RequestSaveFilePicker == null) return;
            var destPath = await RequestSaveFilePicker.Invoke(Path.GetFileNameWithoutExtension(filePath) + imgExtension, imgExtension);
            
            if (destPath == null)
            {
                if (RequestShowAlert != null)
                    await RequestShowAlert.Invoke("알림", "파일 위치 지정을 다시 해주세요.");
                return;
            }

            ImageFilePackageHelper.Package(filePath, destPath, imgPath);

            if (RequestConfirm != null && await RequestConfirm.Invoke("완료", "완료되었습니다! 패키징된 파일의 위치를 탐색기로 열까요?"))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = Path.GetDirectoryName(destPath) ?? "",
                    UseShellExecute = true
                });
            }

            if (PathFile != filePath)
            {
                File.Delete(filePath);
            }
        }
        else if (Directory.Exists(filePath))
        {
            var newFilePath = Path.Combine(Path.GetDirectoryName(filePath) ?? "", Path.GetFileNameWithoutExtension(filePath) + ".zip");
            ImageFilePackageHelper.ZipDir(filePath, newFilePath);

            if (RequestSaveFilePicker == null) return;
            var destPath = await RequestSaveFilePicker.Invoke(Path.GetFileNameWithoutExtension(filePath) + imgExtension, imgExtension);

            if (destPath == null)
            {
                if (RequestShowAlert != null)
                    await RequestShowAlert.Invoke("알림", "파일 위치 지정을 다시 해주세요.");
                return;
            }

            ImageFilePackageHelper.Package(newFilePath, destPath, imgPath);
            File.Delete(newFilePath);

            if (RequestConfirm != null && await RequestConfirm.Invoke("완료", "완료되었습니다! 패키징된 파일의 위치를 탐색기로 열까요?"))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = Path.GetDirectoryName(destPath) ?? "",
                    UseShellExecute = true
                });
            }
        }
        else
        {
            if (RequestShowAlert != null)
            {
                await RequestShowAlert.Invoke("에러",
                    "경로 상에 파일/폴더가 존재하지 않거나 엑세스 권한이 없습니다.\n" +
                    "파일이 존재함에도 에러가 발생한다면 관리자 권한으로 실행하거나, 파일을 다른 위치로 옮긴 후 다시 시도해주세요.");
            }
        }
    }
}