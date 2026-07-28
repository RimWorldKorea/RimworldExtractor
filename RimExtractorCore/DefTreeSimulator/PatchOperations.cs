using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.Extractor;

namespace RimExtractorCore.DefTreeSimulator
{
    /// <summary>
    /// 순수하게 XML 트리를 조작하여 패치를 적용하고,
    /// 조건부 패치(FindMod)를 만날 경우 평행 우주(DefSnapshot)를 분기시킵니다.
    /// </summary>
    public static class PatchOperations
    {
        public static List<DefSnapshot> ApplyPatchRecursive(XElement operationNode, DefSnapshot currentUniverse)
        {
            var resultingUniverses = new List<DefSnapshot> { currentUniverse };
            var opClass = operationNode.Attribute("Class")?.Value;

            switch (opClass)
            {
                case "PatchOperationSequence":
                    var operations = operationNode.Element("operations");
                    if (operations != null)
                    {
                        foreach (var childOp in operations.Elements())
                        {
                            var nextUniverses = new List<DefSnapshot>();
                            // 현재까지 분기된 모든 우주에 대해 다음 패치를 순차적으로 적용
                            foreach (var univ in resultingUniverses)
                            {
                                nextUniverses.AddRange(ApplyPatchRecursive(childOp, univ));
                            }
                            resultingUniverses = nextUniverses;
                        }
                    }
                    break;

                case "PatchOperationFindMod":
                case "JPTools.PatchOperationFindModById":
                    var modsList = operationNode.Element("mods")?.Elements("li").Select(e => e.Value).ToList();
                    var matchNode = operationNode.Element("match");
                    var nomatchNode = operationNode.Element("nomatch");

                    if (modsList != null && modsList.Count > 0)
                    {
                        // 1. 조건이 참일 때의 우주(True Universe) 복제
                        var trueUniverse = currentUniverse.Clone(modsList);
                        var nextUniverses = new List<DefSnapshot>();

                        // 2. True 우주에 match 적용
                        if (matchNode != null)
                        {
                            nextUniverses.AddRange(ApplyPatchRecursive(matchNode, trueUniverse));
                        }
                        else
                        {
                            nextUniverses.Add(trueUniverse);
                        }

                        // 3. 기존 우주(False Universe)에 nomatch 적용
                        if (nomatchNode != null)
                        {
                            nextUniverses.AddRange(ApplyPatchRecursive(nomatchNode, currentUniverse));
                        }
                        else
                        {
                            nextUniverses.Add(currentUniverse); // nomatch가 없으면 원본 유지
                        }

                        resultingUniverses = nextUniverses;
                    }
                    break;

                case "PatchOperationAdd":
                    ApplyAdd(operationNode, currentUniverse.Tree);
                    break;
                case "PatchOperationReplace":
                    ApplyReplace(operationNode, currentUniverse.Tree);
                    break;
                case "PatchOperationInsert":
                    ApplyInsert(operationNode, currentUniverse.Tree);
                    break;
                case "PatchOperationAddModExtension":
                    ApplyAddModExtension(operationNode, currentUniverse.Tree);
                    break;
                case "PatchOperationAttributeAdd":
                    ApplyAttribute(operationNode, currentUniverse.Tree, PatchOperationAttributeMode.Add);
                    break;
                case "PatchOperationAttributeRemove":
                    ApplyAttribute(operationNode, currentUniverse.Tree, PatchOperationAttributeMode.Remove);
                    break;
                case "PatchOperationAttributeSet":
                    ApplyAttribute(operationNode, currentUniverse.Tree, PatchOperationAttributeMode.Set);
                    break;
                default:
                    // Log.Wrn($"지원되지 않는 패치 오퍼레이션: {opClass}");
                    break;
            }

            return resultingUniverses;
        }

        private static void ApplyAdd(XElement curNode, XDocument tree)
        {
            var xpath = curNode.Element("xpath")?.Value;
            var value = curNode.Element("value");
            if (xpath == null || value == null) return;

            var selectNodes = tree.SelectNodesSafe(xpath);
            if (selectNodes == null) return;

            foreach (var selectNode in selectNodes)
            {
                foreach (var valueChildNode in value.Elements())
                {
                    selectNode.Add(new XElement(valueChildNode));
                }
            }
        }

        private static void ApplyReplace(XElement curNode, XDocument tree)
        {
            var xpath = curNode.Element("xpath")?.Value;
            var value = curNode.Element("value");
            if (xpath == null || value == null) return;

            var selectNodes = tree.SelectNodesSafe(xpath)?.ToList();
            if (selectNodes == null) return;

            foreach (var selectNode in selectNodes)
            {
                foreach (var valueChildNode in value.Elements())
                {
                    selectNode.AddBeforeSelf(new XElement(valueChildNode));
                }
                selectNode.Remove();
            }
        }

        private static void ApplyInsert(XElement curNode, XDocument tree)
        {
            var xpath = curNode.Element("xpath")?.Value;
            var value = curNode.Element("value");
            if (xpath == null || value == null) return;

            var selectNodes = tree.SelectNodesSafe(xpath);
            if (selectNodes == null) return;

            foreach (var selectNode in selectNodes)
            {
                var currentTarget = selectNode;
                foreach (var valueChildNode in value.Elements())
                {
                    var newNode = new XElement(valueChildNode);
                    currentTarget.AddAfterSelf(newNode);
                    currentTarget = newNode;
                }
            }
        }

        private static void ApplyAddModExtension(XElement curNode, XDocument tree)
        {
            var xpath = curNode.Element("xpath")?.Value;
            var value = curNode.Element("value");
            if (xpath == null || value == null) return;

            var selectNodes = tree.SelectNodesSafe(xpath);
            if (selectNodes == null) return;

            foreach (var selectNode in selectNodes)
            {
                var modExtensionNode = selectNode.Element("modExtensions");
                if (modExtensionNode == null)
                {
                    modExtensionNode = new XElement("modExtensions");
                    selectNode.Add(modExtensionNode);
                }
                foreach (var valueChildNode in value.Elements())
                {
                    modExtensionNode.Add(new XElement(valueChildNode));
                }
            }
        }

        private static void ApplyAttribute(XElement curNode, XDocument tree, PatchOperationAttributeMode mode)
        {
            var xpath = curNode.Element("xpath")?.Value;
            var value = curNode.Element("value")?.Value;
            var attribute = curNode.Element("attribute")?.Value;

            if (xpath == null || attribute == null) return;
            if (mode != PatchOperationAttributeMode.Remove && value == null) return;

            var selectNodes = tree.SelectNodesSafe(xpath);
            if (selectNodes == null) return;

            foreach (var selectNode in selectNodes)
            {
                switch (mode)
                {
                    case PatchOperationAttributeMode.Add:
                        if (selectNode.Attribute(attribute) == null)
                            selectNode.SetAttributeValue(attribute, value);
                        break;
                    case PatchOperationAttributeMode.Remove:
                        selectNode.Attribute(attribute)?.Remove();
                        break;
                    case PatchOperationAttributeMode.Set:
                        selectNode.SetAttributeValue(attribute, value);
                        break;
                }
            }
        }

        private enum PatchOperationAttributeMode { Add, Remove, Set }
    }
}