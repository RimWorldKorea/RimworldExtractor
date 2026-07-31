using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using RimExtractorCore;
using RimExtractorFace.Views;

namespace RimExtractorFace;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            // Settings.json 환경설정 로드
            SettingManager.Load();
            var config = SettingManager.Current;

            // 경로 문자열이 비어있지 않고, 실제 디스크 상에 존재하는 폴더인지 검증
            bool isPathValid = !string.IsNullOrWhiteSpace(config.PathRimworld) &&
                               !string.IsNullOrWhiteSpace(config.PathWorkshop) &&
                               Directory.Exists(config.PathRimworld) &&
                               Directory.Exists(config.PathWorkshop);

            if (!isPathValid)
            {
                // 경로가 없거나 유효하지 않으면 초기 경로 설정 창 띄우기
                desktop.MainWindow = new InitialPathSelectWindow();
            }
            else
            {
                // 정상적이면 메인 창 띄우기
                desktop.MainWindow = new MainWindow();
            }
        }
        
        base.OnFrameworkInitializationCompleted();
    }
}