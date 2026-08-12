using System.Collections.Concurrent;
using System.Diagnostics;
using System.Xml.Linq;
using ICSharpCode.Decompiler;
using ICSharpCode.Decompiler.CSharp;
using ICSharpCode.Decompiler.CSharp.Syntax;
using ICSharpCode.Decompiler.Metadata;
using ICSharpCode.Decompiler.TypeSystem;

namespace RimExtractorCore.DefTreeSimulator;

public static class AssemblyResolver
{
    public static string GenerateBaseDefTree(string assemblyPath, string outputFilePath)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        
        if (!File.Exists(assemblyPath))
            throw new FileNotFoundException($"어셈블리를 찾을 수 없습니다: {assemblyPath}");

        Log.Msg("IL 어셈블리 분석 및 통합 스키마 트리(PrePiledTree) 생성 중...");

        var decompilerSettings = new DecompilerSettings(LanguageVersion.Latest);
        
        // 메인 스레드용 디컴파일러 (타입 시스템 로드용)
        var mainDecompiler = CreateDecompilerWithResolver(assemblyPath, decompilerSettings);
        var typeSystem = mainDecompiler.TypeSystem;

        var allValidTypes = typeSystem.MainModule.TopLevelTypeDefinitions
            .Where(t => (t.Kind == TypeKind.Class || t.Kind == TypeKind.Struct) &&
                        t.TypeParameterCount == 0 && // 제네릭 타입 차단
                        !t.ReflectionName.StartsWith("System.") &&
                        !t.ReflectionName.StartsWith("UnityEngine.") &&
                        !t.ReflectionName.StartsWith("Unity.") &&
                        !t.ReflectionName.Contains("Steamworks") &&
                        !t.ReflectionName.Contains('<') && 
                        !t.ReflectionName.Contains('>') &&
                        !t.ReflectionName.Contains('$'))   
            .ToList();

        var rootNode = new XElement("PrePiled");
        var defsNode = new XElement("Defs");
        var typesNode = new XElement("Types");

        // [NEW] 스레드 안전한(Thread-Safe) 결과 수집용 컬렉션
        var concurrentDefs = new ConcurrentBag<XElement>();
        var concurrentTypes = new ConcurrentBag<XElement>();
        
        var parallelOptions = new ParallelOptions 
        { 
            /*
             * 5600X 테스트 결과
             * 2 -> 44s
             * 3 -> 37-45s
             * 4 -> 36-41s
             * 5 -> 31-37s
             * 6 -> 31-36s
             * 8 -> 32-34s
             * 그 이상은 유의미한 성능 향상 없었음
             */
            MaxDegreeOfParallelism = Math.Min(8, Environment.ProcessorCount)
        };

        Log.Msg($"어셈블리 내의 모든 유효 타입({allValidTypes.Count}개) 딥 스키마 추출 시작.\n{parallelOptions.MaxDegreeOfParallelism}개 스레드 투입.");

        // [NEW] 스레드별로 독립적인 디컴파일러 인스턴스를 생성하여 할당합니다.
        using (var threadLocalDecompiler = new ThreadLocal<CSharpDecompiler>(() => 
                   CreateDecompilerWithResolver(assemblyPath, decompilerSettings)))
        {
            // [NEW] 병렬 루프 (Parallel.ForEach) 적용
            Parallel.ForEach(allValidTypes, parallelOptions, typeDef =>
            {
                // 공통 메서드 호출!
                var typeNode = ExtractTypeSchemaNode(typeDef, threadLocalDecompiler.Value!);
                
                // 분기 처리만 남음
                bool isDef = InheritsFrom(typeDef, "Verse.Def");
                if (isDef) concurrentDefs.Add(typeNode);
                else concurrentTypes.Add(typeNode);
            });
        }

        // [NEW] 병렬 처리가 끝난 후, 메인 스레드에서 안전하게 한 번에 쏟아 넣습니다.
        defsNode.Add(concurrentDefs.OrderBy(x => x.Name.LocalName));
        typesNode.Add(concurrentTypes.OrderBy(x => x.Name.LocalName));

        rootNode.Add(defsNode, typesNode);
        
        // [수정됨] 누락되었던 XDocument 인스턴스 생성!
        var doc = new XDocument(rootNode);

        doc.Save(outputFilePath);
        Log.Msg($"통합 스키마 추출 완료 (총 {allValidTypes.Count}개 타입 분석됨): {outputFilePath}");
        stopwatch.Stop();
        Log.Msg($"{stopwatch.ElapsedMilliseconds/1000}s 소요");
        
        return outputFilePath;
    }
    
    // [NEW] 여러 모드 어셈블리를 읽어 메모리 상의 typesNode에 병합하는 메서드
    public static void AppendModAssembliesSchema(IEnumerable<string> assemblyPaths, XElement typesNode)
    {
        //TODO 디컴파일러 설정 공부하기
        var decompilerSettings = new DecompilerSettings(LanguageVersion.Latest);
        
        // 현재 시스템의 모든 논리 코어(가용 스레드)를 100% 할당
        var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        foreach (var assemblyPath in assemblyPaths)
        {
            if (!File.Exists(assemblyPath)) continue;

            Log.Msg($"모드 어셈블리 딥 스키마 병합 중... : {Path.GetFileName(assemblyPath)}");

            try
            {
                Log.Common("묑1");
                var mainDecompiler = CreateDecompilerWithResolver(assemblyPath, decompilerSettings);
                Log.Common("묑2");
                var typeSystem = mainDecompiler.TypeSystem;

                
                Log.Common("묑3");
                var allValidTypes = typeSystem.MainModule.TopLevelTypeDefinitions
                    .Where(t => (t.Kind == TypeKind.Class || t.Kind == TypeKind.Struct) &&
                                t.TypeParameterCount == 0 && 
                                !t.ReflectionName.StartsWith("System.") &&
                                !t.ReflectionName.StartsWith("UnityEngine.") &&
                                !t.ReflectionName.StartsWith("Unity.") &&
                                !t.ReflectionName.Contains("Steamworks") &&
                                !t.ReflectionName.Contains('<') && 
                                !t.ReflectionName.Contains('>') &&
                                !t.ReflectionName.Contains('$'))   
                    .ToList();

                var concurrentTypes = new ConcurrentBag<XElement>();

                
                Log.Msg("밍");
                using (var threadLocalDecompiler = new ThreadLocal<CSharpDecompiler>(() => 
                           CreateDecompilerWithResolver(assemblyPath, decompilerSettings)))
                {
                    Parallel.ForEach(allValidTypes, parallelOptions, typeDef =>
                    {
                        Log.Common("핑");
                        // 공통 메서드 호출 후 전부 Types에 담기!
                        var typeNode = ExtractTypeSchemaNode(typeDef, threadLocalDecompiler.Value!);
                        Log.Common("푕");
                        concurrentTypes.Add(typeNode);
                        Log.Common("퐁");
                    });
                }

                // 메모리의 typesNode에 곧바로 병합
                typesNode.Add(concurrentTypes.OrderBy(x => x.Name.LocalName).ToList());
            }
            catch (Exception e)
            {
                Log.Wrn($"어셈블리({assemblyPath}) 분석 중 오류 발생: {e.Message}");
            }
        }
    }
    
    // [NEW] 중복을 제거하기 위해 분리한 "단일 타입 스키마 추출" 핵심 공통 메서드
    private static XElement ExtractTypeSchemaNode(ITypeDefinition typeDef, CSharpDecompiler localDecompiler)
    {
        var typeNode = new XElement(typeDef.Name);

        typeNode.SetAttributeValue("Name", typeDef.Name);
        typeNode.SetAttributeValue("Abstract", "True");

        var baseType = typeDef.DirectBaseTypes.FirstOrDefault(b => 
            b.Kind == TypeKind.Class && 
            b.ReflectionName != "System.Object" &&
            b.ReflectionName != "System.ValueType");

        if (baseType != null)
        {
            typeNode.SetAttributeValue("ParentName", baseType.Name);
        }

        var fieldNodes = new Dictionary<string, XElement>();

        foreach (var field in typeDef.GetFields(f => f.DeclaringTypeDefinition == typeDef && !f.IsConst && !f.IsStatic))
        {
            if (field.Name.Contains('<') || field.Name.Contains('>')) continue;
            if (field.Type.Kind == TypeKind.Unknown || field.Type.ReflectionName.Contains("Steamworks")) continue;
            if (field.GetAttributes().Any(a => a.AttributeType.Name == "UnsavedAttribute")) continue;
            
            var fieldNode = new XElement(field.Name);
            fieldNodes[field.Name] = fieldNode;

            SetTypeInformation(fieldNode, field.Type);

            foreach (var attr in field.GetAttributes())
            {
                if (Constants.TranslationAttributes.TryGetValue(attr.AttributeType.Name, out var mappedName))
                {
                    fieldNode.SetAttributeValue(mappedName, "True");
                }
            }
            
            typeNode.Add(fieldNode);
        }

        try
        {
            var syntaxTree = localDecompiler.DecompileType(typeDef.FullTypeName);
            foreach (var fieldDecl in syntaxTree.Descendants.OfType<FieldDeclaration>())
            {
                foreach (var variable in fieldDecl.Variables)
                {
                    if (variable.Initializer != null && !variable.Initializer.IsNull)
                    {
                        string? valueStr = ParseExpression(variable.Initializer);
                        if (valueStr != null && fieldNodes.TryGetValue(variable.Name, out var fieldNode))
                        {
                            fieldNode.Value = valueStr;
                        }
                    }
                }
            }
        }
        catch (Exception e)
        {
            if (!e.Message.Contains("was not found") && !e.Message.Contains("Unity"))
            {
                Log.Err($"[{typeDef.Name}] 디컴파일 중 예외: {e.Message}");
            }
        }

#if DEBUG
        // 모드 클래스만 필터링해서 확인 (Verse, RimWorld 등 코어 제외)
        if (!typeDef.ReflectionName.StartsWith("Verse.") && !typeDef.ReflectionName.StartsWith("RimWorld."))
        {
            Log.Msg($"[DeepSchema] {typeDef.Name} 딥 스키마 추출 완료 - 부모: {baseType?.Name ?? "없음(Unknown)"}, 추출된 필드: {fieldNodes.Count}개");
        }
#endif
        
        return typeNode;
    }

    private static void SetTypeInformation(XElement fieldNode, IType type)
    {
        // 1. 단일 Enum 처리
        if (type.Kind == TypeKind.Enum)
        {
            fieldNode.SetAttributeValue("Enum", "True");
        }
        
        if (type.Name == "List" && type.TypeArguments.Count == 1)
        {
            fieldNode.SetAttributeValue("List", "True");
            fieldNode.SetAttributeValue("Type", GetFriendlyTypeName(type.TypeArguments[0]));
            return;
        }

        if (type.Name == "Nullable" && type.TypeArguments.Count == 1)
        {
            fieldNode.SetAttributeValue("Type", GetFriendlyTypeName(type.TypeArguments[0]));
            return;
        }

        fieldNode.SetAttributeValue("Type", GetFriendlyTypeName(type));
    }

    private static string GetFriendlyTypeName(IType type)
    {
        return type.ReflectionName switch
        {
            "System.Boolean" => "bool",
            "System.String" => "string",
            "System.Single" => "float",
            "System.Int32" => "int",
            "System.Int64" => "long",
            "System.Double" => "double",
            "System.Byte" => "byte",
            _ => type.Name 
        };
    }

    private static string? ParseExpression(Expression expr)
    {
        if (expr is PrimitiveExpression primitive)
            return primitive.Value?.ToString();

        if (expr is MemberReferenceExpression memberRef)
            return memberRef.MemberName;

        if (expr is ObjectCreateExpression objCreate && objCreate.Arguments.Count == 0)
            return null;

        return null;
    }

    private static bool InheritsFrom(ITypeDefinition? typeDef, string baseTypeName)
    {
        if (typeDef == null) return false;
        if (typeDef.ReflectionName == baseTypeName) return true;
        return typeDef.GetAllBaseTypes().Any(b => b.ReflectionName == baseTypeName);
    }
    
    private static CSharpDecompiler CreateDecompilerWithResolver(string assemblyPath, DecompilerSettings settings)
    {
        var module = new PEFile(assemblyPath);
        var resolver = new UniversalAssemblyResolver(assemblyPath, throwOnError: false, module.DetectTargetFrameworkId());
    
        // 림월드 본편 폴더 강제 주입
        var managedDir = Path.Combine(SettingManager.Current.PathRimworld, "RimWorldWin64_Data", "Managed");
        if (Directory.Exists(managedDir))
        {
            resolver.AddSearchDirectory(managedDir);
        }
    
        return new CSharpDecompiler(module, resolver, settings);
    }
}