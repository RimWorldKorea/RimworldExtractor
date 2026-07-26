using Avalonia.Controls;
using Avalonia.Platform.Storage;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

public partial class ImageFileCombinerWindow : Window
{
    public ImageFileCombinerWindow()
    {
        InitializeComponent();

        var viewModel = new ImageFileCombinerViewModel();
        DataContext = viewModel;

        viewModel.RequestShowAlert += async (title, msg) =>
        {
            var box = MessageBoxManager.GetMessageBoxStandard(title, msg);
            await box.ShowAsync();
        };

        viewModel.RequestConfirm += async (title, msg) =>
        {
            var box = MessageBoxManager.GetMessageBoxStandard(title, msg, ButtonEnum.YesNo);
            return await box.ShowAsync() == ButtonResult.Yes;
        };

        viewModel.RequestSaveFilePicker += async (suggestedName, ext) =>
        {
            var topLevel = GetTopLevel(this);
            if (topLevel == null) return null;

            var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "저장할 파일의 위치를 지정해주세요",
                SuggestedFileName = suggestedName,
                DefaultExtension = ext.TrimStart('.'),
                FileTypeChoices = new[]
                {
                    new FilePickerFileType("이미지 파일") { Patterns = new[] { "*" + ext } }
                }
            });

            return file?.Path.LocalPath;
        };
    }
}