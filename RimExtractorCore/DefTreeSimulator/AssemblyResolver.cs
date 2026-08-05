using System.Collections.Concurrent;
using System.Diagnostics;
using System.Xml.Linq;
using ICSharpCode.Decompiler;
using ICSharpCode.Decompiler.CSharp;
using ICSharpCode.Decompiler.CSharp.Syntax;
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
        var mainDecompiler = new CSharpDecompiler(assemblyPath, decompilerSettings);
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

        Log.Msg($"어셈블리 내의 모든 유효 타입({allValidTypes.Count}개) 딥 스키마 추출 시작 (병렬 처리 중)...");

        // [NEW] 스레드별로 독립적인 디컴파일러 인스턴스를 생성하여 할당합니다.
        using (var threadLocalDecompiler = new ThreadLocal<CSharpDecompiler>(() => 
            new CSharpDecompiler(assemblyPath, decompilerSettings)))
        {
            // [NEW] 병렬 루프 (Parallel.ForEach) 적용
            Parallel.ForEach(allValidTypes, typeDef =>
            {
                var localDecompiler = threadLocalDecompiler.Value!;

                bool isDef = InheritsFrom(typeDef, "Verse.Def");
                var typeNode = new XElement(typeDef.Name);
                
                // [NEW] Def 여부와 상관없이 모든 타입에 Name, Abstract, ParentName을 부여합니다!
                typeNode.SetAttributeValue("Name", typeDef.Name);
                typeNode.SetAttributeValue("Abstract", "True");

                // System.Object나 System.ValueType이 아닌 의미 있는 부모 클래스가 있다면 ParentName 기록
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

                    var transAttrs = field.GetAttributes()
                        .Where(a => a.AttributeType.Name.Contains("Translate"))
                        .ToList();
                    foreach (var attr in transAttrs)
                    {
                        var attrName = attr.AttributeType.Name.Replace("Attribute", "");
                        fieldNode.SetAttributeValue(attrName, "True");
                    }
                    
                    typeNode.Add(fieldNode);
                }

                try
                {
                    // [NEW] 스레드에 할당된 독립적인 디컴파일러를 사용합니다.
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
                    if (!e.Message.Contains("was not found in the module being decompiled") &&
                        !e.Message.Contains("Unity") &&
                        !e.Message.Contains("Steamworks"))
                    {
                        Log.Err($"[{typeDef.Name}] 디컴파일 중 예외: {e.Message}");
                    }
                }

                // [NEW] 안전한 바구니에 담아둡니다.
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

    private static void SetTypeInformation(XElement fieldNode, IType type)
    {
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
}