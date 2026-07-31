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

        Log.Msg("IL 디컴파일러 엔진 초기화");
        
        var decompilerSettings = new DecompilerSettings(LanguageVersion.Latest);
        var decompiler = new CSharpDecompiler(assemblyPath, decompilerSettings);
        var typeSystem = decompiler.TypeSystem;

        var defTypes = typeSystem.MainModule.TopLevelTypeDefinitions
            .Where(t => t.Kind == TypeKind.Class 
                     && !t.IsAbstract 
                     && InheritsFrom(t, "Verse.Def"))
            .ToList();

        Log.Msg($"총 {defTypes.Count}개의 Def 클래스 구조를 XML로 미러링합니다...");

        var rootNode = new XElement("Defs");
        var doc = new XDocument(rootNode);

        foreach (var typeDef in defTypes)
        {
            var defNode = new XElement(typeDef.Name);

            defNode.SetAttributeValue("Name", typeDef.Name); 
            defNode.SetAttributeValue("Abstract", "True");

            var baseType = typeDef.DirectBaseTypes.FirstOrDefault(b => b.Kind == TypeKind.Class && b.ReflectionName != "System.Object");
            if (baseType != null && InheritsFrom(baseType.GetDefinition(), "Verse.Def"))
                defNode.SetAttributeValue("ParentName", baseType.Name);
            
            var fieldNodes = new Dictionary<string, XElement>();

            foreach (var field in typeDef.GetFields(f => f.DeclaringTypeDefinition == typeDef && !f.IsConst && !f.IsStatic))
            {
                if (field.GetAttributes().Any(a => a.AttributeType.Name == "UnsavedAttribute"))
                    continue;

                var fieldNode = new XElement(field.Name);
                fieldNodes[field.Name] = fieldNode;

                // 타입 이름 정제 (List, Nullable, CTS 타입 변환)
                SetTypeInformation(fieldNode, field.Type);

                var transAttrs = field.GetAttributes()
                    .Where(a => a.AttributeType.Name.Contains("Translate"))
                    .ToList();
                foreach (var attr in transAttrs)
                {
                    var attrName = attr.AttributeType.Name.Replace("Attribute", "");
                    fieldNode.SetAttributeValue(attrName, "True");
                }
                
                defNode.Add(fieldNode);
            }

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

            rootNode.Add(defNode);
        }
        
        doc.Save(outputFilePath);
        
        Log.Msg($"사전 모델 트리 생성 완료: {outputFilePath}");
        return outputFilePath;
    }

    /// <summary>
    /// TypeSystem의 타입을 읽어 XML에 예쁘게 기록합니다.
    /// </summary>
    private static void SetTypeInformation(XElement fieldNode, IType type)
    {
        // 1. List 처리
        if (type.Name == "List" && type.TypeArguments.Count == 1)
        {
            fieldNode.SetAttributeValue("List", "True");
            fieldNode.SetAttributeValue("Type", GetFriendlyTypeName(type.TypeArguments[0]));
            return;
        }

        // 2. Nullable (예: int?) 처리
        if (type.Name == "Nullable" && type.TypeArguments.Count == 1)
        {
            fieldNode.SetAttributeValue("Type", GetFriendlyTypeName(type.TypeArguments[0]));
            return;
        }

        // 3. 일반 타입
        fieldNode.SetAttributeValue("Type", GetFriendlyTypeName(type));
    }

    /// <summary>
    /// .NET CTS 타입을 C# 친화적인 이름으로 변환합니다.
    /// </summary>
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
            // 그 외 (Def 이름이나 enum 등)는 그대로 사용 (네임스페이스 제거)
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