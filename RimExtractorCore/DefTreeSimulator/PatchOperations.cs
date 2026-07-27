using System.Xml.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.Extractor;

namespace RimExtractorCore.DefTreeSimulator;

internal static class PatchOperations
{
    public static void ExecutePrePatches(SimulationResult simResult, ExtractableFolder patchDir)
    {
        var patchesRoot = patchDir.FullPath;
        var doc = new XDocument(new XElement("Patch"));
        foreach (var filePath in FileInterface.DescendantFiles(patchesRoot).Where(x => x.ToLower().EndsWith(".xml")))
        {
            var childDoc = FileInterface.ReadXml(filePath);
            foreach (var node in childDoc.Root!.Elements())
            {
                if (node.Name.LocalName != "Operation")
                    continue;

                doc.Root!.Add(new XElement(node));
            }
        }
        
        foreach (var node in doc.Root!.Elements())
        {
            // PatchOperationRecursive가 내부에서 simResult(DefTree, DefsAddedByPatches 등)를
            // 직접 수정하므로, ToList()로 제너레이터를 끝까지 소비시켜 패치 효과를 적용시킵니다.
            _ = PatchOperationRecursive(node, simResult, null, prePatchMode: true).ToList();
        }
    }

    public static IEnumerable<TranslationEntry> PatchOperationRecursive(XElement curNode, SimulationResult simResult, RequiredMods? requiredMods, bool prePatchMode = false)
    {
        if (simResult.DefTree == null)
            yield break;

        var operation = curNode.Attribute("Class")?.Value;

        if (curNode.TryGetAttritube("MayRequire", out string? mayRequire) && mayRequire != null)
        {
            requiredMods = requiredMods != null ? new RequiredMods(requiredMods) : new RequiredMods();
            requiredMods.AddAllowedByPackageIds(mayRequire.Split(','));
        }

        XElement? success = curNode.Element("success");
        switch (operation)
        {
            case "PatchOperationFindMod":
                foreach (var translationEntry in PatchOperationFindMod(curNode, simResult, requiredMods, prePatchMode))
                    yield return translationEntry;
                break;
            case "PatchOperationSequence":
                foreach (var translationEntry in PatchOperationSequence(curNode, simResult, requiredMods, prePatchMode)) 
                    yield return translationEntry;
                break;
            case "PatchOperationAdd":
                foreach (var translationEntry in PatchOperationAdd(curNode, simResult, requiredMods, prePatchMode)) 
                    yield return translationEntry;
                break;
            case "PatchOperationReplace":
                foreach (var translationEntry in PatchOperationReplace(curNode, simResult, requiredMods, prePatchMode))
                    yield return translationEntry;
                break;
            case "PatchOperationAddModExtension":
                foreach (var translationEntry in PatchOperationAddModExtension(curNode, simResult, requiredMods, prePatchMode))
                    yield return translationEntry;
                break;
            case "PatchOperationInsert":
                foreach (var translationEntry in PatchOperationInsert(curNode, simResult, requiredMods, prePatchMode))
                    yield return translationEntry;
                break;
            // PrePatches
            case "PatchOperationAttributeAdd":
                foreach (var translationEntry in PatchOperationAttribute(curNode, simResult, requiredMods, PatchOperationAttributeMode.Add, prePatchMode))
                    yield return translationEntry;
                break;
            case "PatchOperationAttributeRemove":
                foreach (var translationEntry in PatchOperationAttribute(curNode, simResult, requiredMods, PatchOperationAttributeMode.Remove, prePatchMode))
                    yield return translationEntry;
                break;
            case "PatchOperationAttributeSet":
                foreach (var translationEntry in PatchOperationAttribute(curNode, simResult, requiredMods, PatchOperationAttributeMode.Set, prePatchMode))
                    yield return translationEntry;
                break;
            // Added by mods
            case "JPTools.PatchOperationFindModById":
                foreach (var translationEntry in JPTools_PatchOperationFindModById(curNode, simResult, requiredMods, prePatchMode))
                    yield return translationEntry;
                break;
            default:
                Log.Msg($"지원하지 않는 PatchOperation 타입입니다: {operation}");
                break;
        }

        yield break;
    }

    private static IEnumerable<TranslationEntry> PatchOperationInsert(XElement curNode, SimulationResult simResult, RequiredMods? requiredMods, bool prePatchMode = false)
    {
        if (prePatchMode)
            yield break;

        bool isOfficial = simResult.TargetMod?.IsOfficialContent ?? false;
        var xpath = curNode.Element("xpath")?.Value;
        XElement? value = curNode.Element("value");
        if (xpath == null || value == null)
        {
            Log.Wrn($"xpath 또는 value가 없습니다. 잘못된 패치 XML입니다.");
            yield break;
        }

        var selectNodes = simResult.DefTree.SelectNodesSafe(xpath);
        if (selectNodes == null) yield break;
        foreach (XElement selectNode in selectNodes)
        {
            var parentNode = selectNode.Parent;
            if (parentNode == null)
            {
                Log.Wrn($"선택된 노드 {selectNode.Name.LocalName}의 부모 노드가 없습니다.");
                continue;
            }
            var rootDefNode = ExtractorEngine.GetRootDefNode(parentNode, out var nodeName);
            var currentTarget = selectNode;
            foreach (XElement valueChildNode in value.Elements())
            {
                var selectNodeImported = new XElement(valueChildNode);
                currentTarget.AddAfterSelf(selectNodeImported);
                currentTarget = selectNodeImported;

                var curRootDefNode = rootDefNode ?? selectNodeImported;
                var defName = curRootDefNode.Element("defName")?.Value;
                if (defName == null)
                {
                    continue;
                }

                // 🟢 bool isOfficial 인수 추가 반영
                foreach (var translation in ExtractorEngine.FindExtractableNodes(
                             curRootDefNode.Element("defName")!.Value,
                             curRootDefNode.Attribute("Class")?.Value ?? curRootDefNode.Name.LocalName, 
                             selectNodeImported, 
                             isOfficial, 
                             nodeName))
                {
                    yield return translation with
                    {
                        ClassName = $"Patches.{translation.ClassName}",
                        RequiredMods = requiredMods
                    };
                }
            }
        }
    }

    private static IEnumerable<TranslationEntry> PatchOperationAddModExtension(XElement curNode, SimulationResult simResult, RequiredMods? requiredMods, bool prePatchMode = false)
    {
        if (prePatchMode)
            yield break;

        bool isOfficial = simResult.TargetMod?.IsOfficialContent ?? false;
        var xpath = curNode.Element("xpath")?.Value;
        XElement? value = curNode.Element("value");
        if (xpath == null || value == null)
        {
            Log.Wrn($"xpath 또는 value가 없습니다. 잘못된 패치 XML입니다.");
            yield break;
        }

        var selectNodes = simResult.DefTree.SelectNodesSafe(xpath);
        if (selectNodes == null) yield break;
        foreach (XElement selectNode in selectNodes)
        {
            var rootDefNode = ExtractorEngine.GetRootDefNode(selectNode, out var nodeName);
            var modExtensionNode = selectNode.Element("modExtensions");
            if (modExtensionNode == null)
            {
                modExtensionNode = new XElement("modExtensions");
                selectNode.Add(modExtensionNode);
            }

            foreach (XElement valueChildNode in value.Elements())
            {
                var selectNodeImported = new XElement(valueChildNode);
                modExtensionNode.Add(selectNodeImported);
                var curRootDefNode = rootDefNode ?? selectNodeImported;

                // 🟢 bool isOfficial 인수 추가 반영
                foreach (var translation in ExtractorEngine.FindExtractableNodes(
                             curRootDefNode.Element("defName")!.Value,
                             curRootDefNode.Attribute("Class")?.Value ?? curRootDefNode.Name.LocalName, 
                             selectNodeImported, 
                             isOfficial, 
                             nodeName))
                {
                    yield return translation with
                    {
                        ClassName = $"Patches.{translation.ClassName}",
                        RequiredMods = requiredMods
                    };
                }
            }
        }
    }

    private static IEnumerable<TranslationEntry> PatchOperationReplace(XElement curNode, SimulationResult simResult, RequiredMods? requiredMods, bool prePatchMode = false)
    {
        if (prePatchMode)
            yield break;

        bool isOfficial = simResult.TargetMod?.IsOfficialContent ?? false;
        var xpath = curNode.Element("xpath")?.Value;
        XElement? value = curNode.Element("value");
        if (xpath == null || value == null)
        {
            Log.Wrn($"xpath 또는 value가 없습니다. 잘못된 패치 XML입니다.");
            yield break;
        }

        var selectNodes = simResult.DefTree.SelectNodesSafe(xpath);
        if (selectNodes == null) yield break;
        foreach (XElement selectNode in selectNodes)
        {
            var parentNode = selectNode.Parent!;
            var rootDefNode = ExtractorEngine.GetRootDefNode(parentNode, out var nodeName);
            var defName = rootDefNode?.Element("defName")?.Value;
            var className = (rootDefNode?.Attribute("Class")?.Value ?? rootDefNode?.Name.LocalName);
                            
            if (rootDefNode == null)
            {
                defName = selectNode.Element("defName")?.Value;
                className = (selectNode.Attribute("Class")?.Value ?? selectNode.Name.LocalName);
            }
            if (defName is null || className is null)
                Log.Wrn($"defName 또는 className을 찾을 수 없는 Patch: xpath:{xpath}");
            
            var currentTarget = selectNode;
            foreach (XElement valueChildNode in value.Elements())
            {
                var selectNodeImported = new XElement(valueChildNode);
                currentTarget.AddBeforeSelf(selectNodeImported);

                // 🟢 bool isOfficial 인수 추가 반영
                foreach (var translation in ExtractorEngine.FindExtractableNodes(defName, className, selectNodeImported, isOfficial, nodeName))
                {
                    yield return translation with
                    {
                        ClassName = $"Patches.{translation.ClassName}",
                        RequiredMods = requiredMods
                    };
                }
            }
            selectNode.Remove();
        }
    }

    private static IEnumerable<TranslationEntry> PatchOperationAdd(XElement curNode, SimulationResult simResult, RequiredMods? requiredMods, bool prePatchMode = false) 
    {
        if (prePatchMode)
            yield break;

        bool isOfficial = simResult.TargetMod?.IsOfficialContent ?? false;
        var xpath = curNode.Element("xpath")?.Value;
        XElement? value = curNode.Element("value");
        if (xpath == null || value == null)
        {
            Log.Wrn($"xpath 또는 value가 없습니다. 잘못된 패치 XML입니다.");
            yield break;
        }

        var selectNodes = simResult.DefTree.SelectNodesSafe(xpath);
        if (selectNodes == null) yield break;
        foreach (XElement selectNode in selectNodes)
        {
            var rootDefNode = ExtractorEngine.GetRootDefNode(selectNode, out var nodeName);
            foreach (XElement valueChildNode in value.Elements())
            {
                var selectNodeImported = new XElement(valueChildNode);
                selectNode.Add(selectNodeImported);
                var curRootDefNode = rootDefNode ?? selectNodeImported;

                if (xpath is "Defs" or "Defs/")
                {
                    simResult.DefsAddedByPatches.Add((requiredMods, selectNodeImported));
                    continue;
                }

                var defName = curRootDefNode.Element("defName")?.Value;
                if (defName == null)
                {
                    Log.Wrn($"defName이 없는 노드입니다. xpath={xpath}, value={value}");
                    continue;
                }

                // 🟢 bool isOfficial 인수 추가 반영
                foreach (var translation in ExtractorEngine.FindExtractableNodes(
                             defName,
                             curRootDefNode.Attribute("Class")?.Value ?? curRootDefNode.Name.LocalName, 
                             selectNodeImported, 
                             isOfficial, 
                             nodeName))
                {
                    if (translation.ClassName == "Keyed")
                    {
                        yield return translation with
                        {
                            RequiredMods = requiredMods
                        };
                        continue;
                    }
                    yield return translation with
                    {
                        ClassName = $"Patches.{translation.ClassName}",
                        RequiredMods = requiredMods
                    };
                }
            }
        }
    }

    private static IEnumerable<TranslationEntry> PatchOperationSequence(XElement curNode, SimulationResult simResult, RequiredMods? requiredMods, bool prePatchMode = false)
    {
        var operations = curNode.Element("operations");
        if (operations == null)
            yield break;
        foreach (XElement childOperation in operations.Elements())
        {
            foreach (var translationEntry in PatchOperationRecursive(childOperation, simResult, requiredMods, prePatchMode))
            {
                yield return translationEntry;
            }
        }
    }

    private static IEnumerable<TranslationEntry> PatchOperationFindMod(XElement curNode, SimulationResult simResult, RequiredMods? requiredMods, bool prePatchMode = false)
    {
        requiredMods = new RequiredMods(requiredMods);
        var noMatchRequiredMods = new RequiredMods(requiredMods);
        var requiredModNodes = curNode.Element("mods")?.Elements();
        var requiredModsList = requiredModNodes?.Select(n => n.Value).ToList();

        var match = curNode.Element("match");
        if (match != null)
        {
            if (requiredModsList != null)
            {
                requiredMods.AddAllowedByModNames(requiredModsList);
            }
            foreach (var translationEntry in PatchOperationRecursive(match, simResult, requiredMods, prePatchMode))
            {
                yield return translationEntry;
            }
        }

        var noMatch = curNode.Element("nomatch");
        if (noMatch != null)
        {
            if (requiredModsList != null)
            {
                noMatchRequiredMods.AddDisallowedByModNames(requiredModsList);
            }
            foreach (var translationEntry in PatchOperationRecursive(noMatch, simResult, noMatchRequiredMods, prePatchMode))
            {
                yield return translationEntry;
            }
        }
    }

    private static IEnumerable<TranslationEntry> PatchOperationAttribute(XElement curNode, SimulationResult simResult, RequiredMods? requiredMods, PatchOperationAttributeMode mode, bool prePatchMode = false)
    {
        var xpath = curNode.Element("xpath")?.Value;
        var value = curNode.Element("value")?.Value;
        var attribute = curNode.Element("attribute")?.Value;
        if (xpath == null || (mode != PatchOperationAttributeMode.Remove && value == null) || attribute == null)
        {
            Log.Wrn($"xpath 또는 value가 없습니다. 잘못된 패치 XML입니다.");
            yield break;
        }

        var selectNodes = simResult.DefTree.SelectNodesSafe(xpath);
        if (selectNodes == null) yield break;
        foreach (XElement selectNode in selectNodes)
        {
            switch (mode)
            {
                case PatchOperationAttributeMode.Add:
                    if (selectNode.Attribute(attribute) == null)
                    {
                        selectNode.SetAttributeValue(attribute, value);
                    }
                    break;
                case PatchOperationAttributeMode.Remove:
                    selectNode.Attribute(attribute)?.Remove();
                    break;
                case PatchOperationAttributeMode.Set:
                    selectNode.SetAttributeValue(attribute, value);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(mode), mode, null);
            }
        }
    }

    private enum PatchOperationAttributeMode
    {
        Add, Remove, Set
    }

    private static IEnumerable<TranslationEntry> JPTools_PatchOperationFindModById(XElement curNode, SimulationResult simResult, RequiredMods? requiredMods, bool prePatchMode = false)
    {
        requiredMods = new RequiredMods(requiredMods);
        var noMatchRequiredMods = new RequiredMods(requiredMods);
        var requiredModNodes = curNode.Element("mods")?.Elements();
        var requiredModsList = requiredModNodes?.Select(n => n.Value).ToList();

        var match = curNode.Element("match");
        if (match != null)
        {
            if (requiredModsList != null)
            {
                requiredMods.AddAllowedByPackageIds(requiredModsList);
            }
            foreach (var translationEntry in PatchOperationRecursive(match, simResult, requiredMods, prePatchMode))
            {
                yield return translationEntry;
            }
        }

        var noMatch = curNode.Element("nomatch");
        if (noMatch != null)
        {
            if (requiredModsList != null)
            {
                noMatchRequiredMods.AddAllowedByPackageIds(requiredModsList);
            }
            foreach (var translationEntry in PatchOperationRecursive(noMatch, simResult, noMatchRequiredMods, prePatchMode))
            {
                yield return translationEntry;
            }
        }
    }
}