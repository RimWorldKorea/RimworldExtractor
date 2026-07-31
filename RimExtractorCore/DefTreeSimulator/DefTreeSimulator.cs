using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.Procedures;

namespace RimExtractorCore.DefTreeSimulator
{
    /// <summary>
    /// 림월드 엔진의 XML 로딩, 패치(PatchOperation), 상속(Inheritance) 과정을 모사하여
    /// 추출기가 분석할 수 있는 최종 런타임 형태의 가상 DOM(DefSnapshot 리스트)을 생성합니다.
    /// </summary>
    public static class DefTreeSimulator
    {
        public static SimulationResult Execute(
            ModMetadata modMetadata,
            List<ExtractableFolder> selectedFolders, 
            List<ExtractableFolder> prePatches,
            List<string>? referenceDefsRoots,
            bool isOfficialContent,
            List<ModMetadata>? referenceMods = null)
        {
            var result = new SimulationResult { TargetMod = modMetadata };

            Log.Msg("1. 사전 어셈블리 모델(PrePiledTree) 로드...");
            XDocument prePiledTree = LoadOrGeneratePrePiledTree();

            // -------------------------------------------------------------------------
            // 2. LoadFolders.xml 조건(Any/All)에 따른 초기 스냅샷 멀티버스 생성
            Log.Msg("2. LoadFolders 조건(Any/All)에 따른 초기 스냅샷 분기 생성...");
            List<DefSnapshot> multiverse = DoctorStrange.GenerateInitialSnapshots(prePiledTree, selectedFolders);

            // -------------------------------------------------------------------------
            // 3. 각 스냅샷별로 자신에게 할당된 Defs 병합
            Log.Msg("3. 분기된 스냅샷별 Defs XML 로드 및 병합...");
            foreach (var snapshot in multiverse)
            {
                LoadAndMergeModDefs(snapshot.Tree, snapshot.AssignedFolders);
            }

            // -------------------------------------------------------------------------
            // 4. 각 스냅샷별 PatchOperation 분기 검사 및 추가 분열
            Log.Msg("4. PatchOperationFindMod 등 조건부 패치를 통한 스냅샷 2차 분열...");
            List<DefSnapshot> finalMultiverse = new List<DefSnapshot>();
            
            foreach (var snapshot in multiverse)
            {
                var branchedSnapshots = DoctorStrange.ProcessPatchOperations(snapshot, snapshot.AssignedPatches);
                finalMultiverse.AddRange(branchedSnapshots);
            }

            // -------------------------------------------------------------------------
            // 5. 상속(Inheritance) 처리 (모든 최종 평행 우주에 대해 각각 수행)
            Log.Msg("5. 최종 생성된 모든 스냅샷에 대해 XML 상속(ParentName) 처리...");
            foreach (var snapshot in finalMultiverse)
            {
                DoXmlInheritance(snapshot.Tree.Root);
                CleanUpAbstractNodes(snapshot.Tree.Root);
            }

            // -------------------------------------------------------------------------
            // 6. PostProcessor (Stage A 등)
            foreach (var snapshot in finalMultiverse)
            {
                snapshot.Tree = XDocumentProcedureInjector.Instance.ExecuteStage(InjectionStage.StageA, snapshot.Tree);
            }

            result.Snapshots = finalMultiverse;
            return result;
        }

        /// <summary>
        /// 어셈블리로부터 C# 클래스 구조를 미러링한 뼈대 트리를 가져옵니다.
        /// </summary>
        private static XDocument LoadOrGeneratePrePiledTree()
        {
            var path = ExtractorCore.PrePiledTreePath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                throw new FileNotFoundException($"사전 트리 XML 파일을 찾을 수 없습니다: {path}");
            }
            return XDocument.Load(path);
        }

        /// <summary>
        /// 추출된 Defs 폴더들의 XML을 읽어 Base 트리에 병합합니다.
        /// </summary>
        private static void LoadAndMergeModDefs(XDocument baseTree, List<ExtractableFolder> defFolders)
        {
            var baseRoot = baseTree.Root;
            if (baseRoot == null) return;

            foreach (var folder in defFolders)
            {
                var folderPath = folder.FullPath;
                if (!Directory.Exists(folderPath)) continue;

                // 해당 Defs 폴더 안의 모든 .xml 파일을 탐색합니다.
                var xmlFiles = FileInterface.DescendantFiles(folderPath)
                    .Where(x => x.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        // FileInterface.ReadXml을 사용하여 주석/공백을 무시하고 안전하게 로드합니다.
                        var modDefDoc = FileInterface.ReadXml(xmlFile);
                        var modRoot = modDefDoc.Root;

                        // 루트가 <Defs>인 경우에만 그 안의 자식 노드들을 베이스 트리에 병합합니다.
                        if (modRoot != null && modRoot.Name.LocalName == "Defs")
                        {
                            // 요소들을 베이스 트리의 Root에 추가 (LINQ to XML이 자동으로 복제/이동 처리)
                            baseRoot.Add(modRoot.Elements());
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Wrn($"Def XML 파일 병합 중 오류 발생 ({xmlFile}): {e.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// XML 트리의 ParentName을 추적하여 상속(Inheritance) 병합을 수행합니다.
        /// </summary>
        private static void DoXmlInheritance(XElement? root)
        {
            if (root == null) return;

            var allDefs = root.Elements().ToList();
            var resolvedSet = new HashSet<XElement>();
            
            var defsByName = allDefs
                .Where(e => e.Attribute("Name") != null)
                .ToDictionary(e => e.Attribute("Name")!.Value, e => e);

            foreach (var def in allDefs)
            {
                ResolveInheritanceRecursive(def, defsByName, resolvedSet);
            }
        }

        private static void ResolveInheritanceRecursive(XElement def, Dictionary<string, XElement> defsByName, HashSet<XElement> resolvedSet)
        {
            if (resolvedSet.Contains(def)) return;

            var parentNameAttr = def.Attribute("ParentName");
            if (parentNameAttr != null)
            {
                string parentName = parentNameAttr.Value;
                if (defsByName.TryGetValue(parentName, out var parentDef))
                {
                    ResolveInheritanceRecursive(parentDef, defsByName, resolvedSet);
                    CopyMissingElements(parentDef, def);
                }
            }
            
            resolvedSet.Add(def);
        }

        private static void CopyMissingElements(XElement parent, XElement child)
        {
            foreach (var attr in parent.Attributes())
            {
                if (attr.Name == "Name" || attr.Name == "Abstract" || attr.Name == "ParentName") continue;
                if (child.Attribute(attr.Name) == null)
                {
                    child.Add(new XAttribute(attr));
                }
            }

            foreach (var parentElem in parent.Elements())
            {
                var childElem = child.Element(parentElem.Name);
                if (childElem == null)
                {
                    child.Add(new XElement(parentElem));
                }
                else
                {
                    CopyMissingElements(parentElem, childElem);
                }
            }
        }

        /// <summary>
        /// 상속 처리가 모두 끝난 후, 불필요한 Abstract="True" 껍데기 노드들을 삭제합니다.
        /// </summary>
        private static void CleanUpAbstractNodes(XElement? root)
        {
            if (root == null) return;

            var abstractNodes = root.Elements()
                .Where(e => e.Attribute("Abstract")?.Value.ToLower() == "true")
                .ToList();

            foreach (var node in abstractNodes)
            {
                node.Remove();
            }
        }
    }
}