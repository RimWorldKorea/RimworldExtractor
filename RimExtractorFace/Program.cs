using Avalonia;
using System.Reflection;
using RimExtractorCore; // 🟢 Internal 네임스페이스 추가

namespace RimExtractorFace;

internal static class Program
{
    public const string VERSION = ""; // Github Action에 의해 게시 전 자동으로 생성

    [STAThread]
    public static void Main(string[] args)
    {
        AppDomain.CurrentDomain.AssemblyResolve += CurrentDomainOnAssemblyResolve;

        // 🟢 메인 윈도우 렌더링을 방해하지 않도록 백그라운드 스레드에서 프로시저 동적 컴파일 시작
        Task.Run(() =>
        {
            try
            {
                Engine.Initialize();
            }
            catch (Exception ex)
            {
                Log.Err($"[GUI] 초기화 중 예외 발생: {ex.Message}");
            }
        });

        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();

    // 기존 dll 동적 로드 로직 유지
    private static Assembly? CurrentDomainOnAssemblyResolve(object? sender, ResolveEventArgs args)
    {
        if (args.Name.Contains(".resources")) return null;

        Assembly? assembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.FullName == args.Name);
        if (assembly != null) return assembly;

        string filename = args.Name.Split(',')[0] + ".dll".ToLower();
        var assemblyFilePath = Path.Combine("bin", filename);

        if (File.Exists(assemblyFilePath))
        {
            try
            {
                return Assembly.LoadFrom(assemblyFilePath);
            }
            catch
            {
                return null;
            }
        }
        return null;
    }
}