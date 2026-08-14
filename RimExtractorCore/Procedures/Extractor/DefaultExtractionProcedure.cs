using System.Collections;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.Procedures;

public class DefaultExtractionProcedure : IExtractionProcedure
{
    public string Name => "DefaultExtractionProcedure";

    public IEnumerable<TranslationEntry> Extract(DefSnapshot snapshot, ModMetadata targetMod)
    {
        if (snapshot.Tree.Root == null) yield break;

        bool isOfficialContent = targetMod.IsOfficialContent;

        // Reference 노드는 제외하고, ExtractionTarget 마커가 있는 모드 고유 노드만 추출 대상에 포함
        var rootElements = snapshot.Tree.Root.Elements()
            .Where(x => x.Attribute("Reference")?.Value.ToLower() != "true" &&
                        x.Attribute("ExtractionTarget")?.Value == "True");

        foreach (var defNode in rootElements)
        {
            var defName = defNode.Element("defName")?.Value;
            if (string.IsNullOrEmpty(defName)) continue;

            var className = defNode.Attribute("Class")?.Value ?? defNode.Name.LocalName;
            className = className[..1].ToUpper() + className[1..];

            // 트리를 순회하며 Type="string" 인 텍스트 노드 추출
            foreach (var entry in TraverseAndExtract(defName, className, defNode, defNode.Name.LocalName,
                         isOfficialContent))
            {
                yield return entry;
            }
        }
    }

    private IEnumerable<TranslationEntry> TraverseAndExtract(
        string defName, string className, XElement curNode, string currentPath, bool isOfficialContent)
    {
        foreach (var child in curNode.Elements())
        {
            // 명시적인 NoTranslate 어트리뷰트가 있는 경우 스킵합니다.
            //TODO 굳이 변수로 단계를 나눌 필요가 있는진 모르겠음.
            bool isNoTranslate = child.Attribute(Constants.AttrNoTranslate)?.Value.ToLower() == "true";
            bool isFullListTranslate = false;
            
            if (isNoTranslate) continue;
            
            string path;
            
            // 리스트 노드(li) 처리
            if (child.Name.LocalName == "li")
            {
                int index = child.ElementsBeforeSelf("li").Count();
                path = $"{currentPath}.{index}";
            }
            else
            {
                path = $"{currentPath}.{child.Name.LocalName}";
            }

            // 자식이 또 있는 컨테이너 노드라면 재귀 탐색
            if (child.HasElements)
            {
                foreach (var entry in TraverseAndExtract(defName, className, child, path, isOfficialContent))
                {
                    yield return entry;
                }
            }
            // 자식이 없는 리프 노드인 경우 (값이 비어있더라도 뼈대 생성을 위해 무조건 진입)
            else
            {
                bool isDefName = child.Name.LocalName == "defName"; // defName은 번역 대상이 아니므로 고정 제외

                // MayNotTranslate 확인
                bool mayNotTranslate = child.Attribute(Constants.AttrMayNotTranslate)?.Value.ToLower() == "true";
                
                // 부모가 List면 부모의 MayNotTranslate도 확인
                if (child.Parent != null && child.Parent.Attribute("List")?.Value == "True")
                {
                    mayNotTranslate =
                        mayNotTranslate ||
                        child.Parent.Attribute(Constants.AttrMayNotTranslate)?.Value.ToLower() == "true";
                }

                // 추출기 설정에 따라 비필수 노드 추출 스킵
                if (mayNotTranslate && !SettingManager.Current.ExtractMayNotTranslate) continue; 
                
                
                // 1. Type 어트리뷰트가 아예 없거나
                // 2. string인 것만 추출
                string? typeAttr = child.Attribute("Type")?.Value;

                // Enum="True"인지 검사합니다
                if (child.Parent != null && child.Parent.Attribute("Enum")?.Value == "True")
                {
                    typeAttr = child.Parent.Attribute("Type")?.Value;
                }
                
                // 부모가 List면 부모의 타입을 가져옵니다.
                if (child.Parent != null && child.Parent.Attribute("List")?.Value == "True")
                {
                    typeAttr = child.Parent.Attribute("Type")?.Value;

                    // 부모가 FullListTranslate면 속성을 가져옵니다.
                    if (child.Parent.Attribute(Constants.AttrTranslationCanChangeCount)?.Value.ToLower() == "true")
                        isFullListTranslate = true;
                }

                bool isStringOrUntyped = string.IsNullOrEmpty(typeAttr) || typeAttr == "string";

                // NoTranslate가 아니고, defName이 아닌 문자열(또는 타입 불명) 노드 추출
                if (isStringOrUntyped && !isNoTranslate && !isDefName)
                {
                    var rootDefNode = curNode.AncestorsAndSelf().LastOrDefault();
                    var requiredModsInnerText = rootDefNode?.Element("REQUIREDMODS")?.Value;
                    var requiredMods = requiredModsInnerText != null
                        ? RequiredMods.FromStringByModNames(requiredModsInnerText)
                        : null;
                    var fileName = isOfficialContent ? rootDefNode?.Attribute("SourceFile")?.Value : null;

                    // currentPath가 "ThingDef.label" 형태이므로 첫 번째 요소를 제외하고 defName을 붙임
                    string nodePath = $"{defName}.{path.Substring(path.IndexOf('.') + 1)}";

                    // 값이 비어있어도(child.Value == "") 뼈대 생성을 위해 그대로 반환
                    yield return new TranslationEntry(
                        className,
                        nodePath,
                        child.Value,
                        null,
                        requiredMods,
                        fileName
                    )
                    {
                        MayNotTranslate = mayNotTranslate,
                        FullListTranslate = isFullListTranslate
                    };
                }
            }
        }
    }
}