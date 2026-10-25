using Avalonia;
using System.Reflection;
using System.Reflection.Metadata;
using RimExtractorCore;

namespace RimExtractorFace;

internal static class Program
{
    //TODO 이거 기존엔 GitHub Action으로 자동 빌드하면서 넣어줬던 것 같은데 어떻게 할지 나중에 결정
    public static readonly string VERSION =
        Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion ?? String.Empty;

    [STAThread]
    public static void Main(string[] args)
    {
        AppDomain.CurrentDomain.AssemblyResolve += CurrentDomainOnAssemblyResolve;
        
        Task.Run(() =>
        {
            try
            {
                Engine.Initialize();
            }
            catch (Exception ex)
            {
                Log.Err($"초기화 중 예외 발생\n\t{ex.Message}");
            }
        });

        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
    
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