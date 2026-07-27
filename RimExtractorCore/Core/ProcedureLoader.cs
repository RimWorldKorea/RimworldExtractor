using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace RimExtractorCore;

public static class ProcedureLoader
{
    /// <summary>
    /// 지정된 경로의 .cs 파일들을 런타임에 컴파일하여 TInterface를 구현하는 인스턴스 목록을 생성합니다.
    /// </summary>
    /// <typeparam name="TInterface">이 인터페이스를 상속하는 프로시저만 불러옵니다.</typeparam>
    /// <param name="folderPath">.cs 파일들이 위치한 폴더 경로입니다.</param>
    public static List<TInterface> LoadProceduresFromDirectory<TInterface>(string folderPath) where TInterface : class
    {
        var result = new List<TInterface>();

        if (!Directory.Exists(folderPath))
        {
            Directory.CreateDirectory(folderPath);
            return result;
        }

        var scriptFiles = Directory.GetFiles(folderPath, "*.cs", SearchOption.AllDirectories);
        foreach (var scriptFile in scriptFiles)
        {
            try
            {
                var instance = CompileAndCreateInstance<TInterface>(scriptFile);
                if (instance != null)
                {
                    result.Add(instance);
                    //Log.Msg($"스크립트 로드 성공 ({Path.GetFileName(folderPath)}): {Path.GetFileName(scriptFile)}");
                }
            }
            catch (Exception e)
            {
                Log.Err($"{scriptFile} 로드 실패\n{e.Message}");
            }
        }

        return result;
    }

    /// <summary>
    /// .cs 파일을 컴파일하여 TInterface 타입의 인스턴스를 생성합니다.
    /// </summary>
    public static TInterface? CompileAndCreateInstance<TInterface>(string filePath) where TInterface : class
    {
        var code = File.ReadAllText(filePath);
        var syntaxTree = CSharpSyntaxTree.ParseText(code);

        var assemblyName = Path.GetRandomFileName();
        var targetType = typeof(TInterface);

        var references = new List<MetadataReference>
        {
            MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(Enumerable).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(System.Xml.Linq.XDocument).Assembly.Location),
            MetadataReference.CreateFromFile(targetType.Assembly.Location)
            
        };

        // 로드된 현재 AppDomain의 비동적 어셈블리 참조 추가
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies().Where(a => !a.IsDynamic && !string.IsNullOrEmpty(a.Location)))
        {
            references.Add(MetadataReference.CreateFromFile(asm.Location));
        }

        var compilation = CSharpCompilation.Create(
            assemblyName,
            syntaxTrees: new[] { syntaxTree },
            references: references,
            options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        using var ms = new MemoryStream();
        var emitResult = compilation.Emit(ms);

        if (!emitResult.Success)
        {
            var failures = emitResult.Diagnostics.Where(diagnostic =>
                diagnostic.IsWarningAsError || diagnostic.Severity == DiagnosticSeverity.Error);

            var errorMsg = string.Join("\n", failures.Select(f => $"{f.Id}: {f.GetMessage()}"));
            throw new InvalidOperationException($"컴파일 오류:\n{errorMsg}");
        }

        ms.Seek(0, SeekOrigin.Begin);
        var assembly = Assembly.Load(ms.ToArray());

        var implementationType = assembly.GetTypes()
            .FirstOrDefault(t => targetType.IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract);

        if (implementationType == null)
        {
            throw new InvalidOperationException($"{targetType.Name} 인터페이스를 구현한 클래스를 찾을 수 없습니다.");
        }

        return (TInterface?)Activator.CreateInstance(implementationType);
    }
}