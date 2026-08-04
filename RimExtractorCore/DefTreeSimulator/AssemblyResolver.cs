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
        if (!File.Exists(assemblyPath))
            throw new FileNotFoundException($"어셈블리를 찾을 수 없습니다: {assemblyPath}");

        Log.Msg("IL 어셈블리 분석 및 통합 스키마 트리(PrePiledTree) 생성 중...");

        var decompilerSettings = new DecompilerSettings(LanguageVersion.Latest);
        var decompiler = new CSharpDecompiler(assemblyPath, decompilerSettings);
        var typeSystem = decompiler.TypeSystem;

        var defTypes = typeSystem.MainModule.TopLevelTypeDefinitions
            .Where(t => t.Kind == TypeKind.Class && InheritsFrom(t, "Verse.Def"))
            .ToList();

        // [NEW] 루트 노드 구조 변경: <PrePiled> 하위에 <Defs>와 <Types>를 병렬로 배치
        var rootNode = new XElement("PrePiled");
        var defsNode = new XElement("Defs");
        var typesNode = new XElement("Types");
        rootNode.Add(defsNode, typesNode);

        var doc = new XDocument(rootNode);

        // [NEW] 순환 참조 방지 및 재귀 탐색을 위한 Queue와 HashSet
        var processedTypes = new HashSet<string>();
        var typeQueue = new Queue<ITypeDefinition>();

        // 1. 모든 Def 타입들을 먼저 큐에 넣습니다.
        foreach (var typeDef in defTypes)
        {
            typeQueue.Enqueue(typeDef);
            processedTypes.Add(typeDef.ReflectionName);
        }

        Log.Msg($"발견된 Def 타입 및 파생 복합 타입(Deep Schema) 추출 시작...");

        // 2. 큐가 빌 때까지 모든 복합 타입들의 필드를 파고듭니다.
        while (typeQueue.Count > 0)
        {
            var typeDef = typeQueue.Dequeue();
            bool isDef = InheritsFrom(typeDef, "Verse.Def");

            var typeNode = new XElement(typeDef.Name);
            
            // Def인 경우에만 Name, Abstract, ParentName을 세팅합니다.
            if (isDef)
            {
                typeNode.SetAttributeValue("Name", typeDef.Name);
                typeNode.SetAttributeValue("Abstract", "True");

                var baseType = typeDef.DirectBaseTypes.FirstOrDefault(b => b.Kind == TypeKind.Class && b.ReflectionName != "System.Object");
                if (baseType != null && InheritsFrom(baseType.GetDefinition(), "Verse.Def"))
                    typeNode.SetAttributeValue("ParentName", baseType.Name);
            }

            var fieldNodes = new Dictionary<string, XElement>();
            
            foreach (var field in typeDef.GetFields(f =>
                         f.DeclaringTypeDefinition == typeDef && !f.IsConst && !f.IsStatic))
            {
                try
                {
                    // 컴파일러가 자동 생성한 Backing Field (이름에 '<' 포함)는 무조건 스킵합니다!
                    if (field.Name.Contains('<') || field.Name.Contains('>'))
                        continue;
                    
                    // 해석할 수 없거나 스팀 관련 타입이므로, 예외를 발생시키기 전에 안전하게 스킵합니다.
                    if (field.Type.Kind == TypeKind.Unknown || field.Type.ReflectionName.Contains("Steamworks"))
                        continue;
                    
                    // 해당 타입의 모든 필드 조사
                    if (field.GetAttributes().Any(a => a.AttributeType.Name == "UnsavedAttribute"))
                        continue;

                    var fieldNode = new XElement(field.Name);
                    fieldNodes[field.Name] = fieldNode;

                    // 기존 타입 정보 기록
                    SetTypeInformation(fieldNode, field.Type);

                    var transAttrs = field.GetAttributes()
                        .Where(a => a.AttributeType.Name.Contains("Translate"))
                        .ToList();
                    foreach (var attr in transAttrs)
                    {
                        var attrName = attr.AttributeType.Name.Replace("Attribute", "");
                        fieldNode.SetAttributeValue(attrName, "True");
                    }

                    // [NEW] 이 필드의 타입이 복합 타입(Class/Struct)이라면 큐에 추가하여 딥 스키마를 추적합니다.
                    EnqueueIfComplex(field.Type, typeQueue, processedTypes);

                    typeNode.Add(fieldNode);

                }
                catch (Exception e)
                {
                    Log.Err(e.Message);
                    Log.Err(field.Name);
                    Log.Err(field.ReflectionName);
                }
            }

            try
            {
                // 구문 트리를 디컴파일하여 기본값(Default Value)을 채워 넣습니다.
                var syntaxTree = decompiler.DecompileType(typeDef.FullTypeName);
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
                Log.Err(e.Message);
            }


            // Def는 <Defs>에, 그 외의 커스텀 파생 클래스(GraphicData 등)는 <Types>에 저장합니다.
            if (isDef) defsNode.Add(typeNode);
            else typesNode.Add(typeNode);
        }
        
        Log.Msg("쮸삣쮸삣");

        doc.Save(outputFilePath);

        Log.Msg($"\n통합 스키마 추출 완료 (총 {processedTypes.Count}개 타입 분석됨): {outputFilePath}");
        return outputFilePath;
    }

    /// <summary>
    /// 복합 타입(Class, Struct)일 경우 딥 스키마 추출을 위해 큐에 등록합니다.
    /// </summary>
    private static void EnqueueIfComplex(IType type, Queue<ITypeDefinition> queue, HashSet<string> processedTypes)
    {
        // 1. List, Nullable, Dictionary 등 제네릭 언래핑 (Unwrapping)
        if (type.Name == "List" && type.TypeArguments.Count == 1)
        {
            EnqueueIfComplex(type.TypeArguments[0], queue, processedTypes);
            return;
        }
        if (type.Name == "Nullable" && type.TypeArguments.Count == 1)
        {
            EnqueueIfComplex(type.TypeArguments[0], queue, processedTypes);
            return;
        }
        if (type.Name == "Dictionary" && type.TypeArguments.Count == 2)
        {
            // Dictionary는 주로 Value 쪽에 복합 타입이 들어갑니다 (예: Dictionary<string, GraphicData>)
            EnqueueIfComplex(type.TypeArguments[1], queue, processedTypes);
            return;
        }

        var def = type.GetDefinition();
        if (def != null)
        {
            // Enum은 속성을 가지지 않으므로 제외, System 및 UnityEngine 기본 구조체/클래스도 딥 탐색에서 제외합니다.
            if (def.Kind != TypeKind.Enum && 
                !def.ReflectionName.StartsWith("System.") && 
                !def.ReflectionName.StartsWith("UnityEngine.") &&
                !def.ReflectionName.StartsWith("Unity."))
            {
                if (!processedTypes.Contains(def.ReflectionName))
                {
                    processedTypes.Add(def.ReflectionName);
                    queue.Enqueue(def);
                }
            }
        }
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