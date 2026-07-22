using Avalonia;
using System.Reflection;

namespace RimworldExtractorGUI;

internal static class Program
{
    public const string VERSION = ""; // Github Action에 의해 게시 전 자동으로 생성

    [STAThread]
    public static void Main(string[] args)
    {
        AppDomain.CurrentDomain.AssemblyResolve += CurrentDomainOnAssemblyResolve;
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();

    // 기존 dll 동적 로드 로직 유지[cite: 1]
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