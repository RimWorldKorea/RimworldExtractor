using System;
using System.Collections.Generic;
using System.Diagnostics;
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


            Log.Msg("Phase1/ 사전 어셈블리 모델(PrePiledTree) 로드...");
            XDocument prePiledTree = LoadOrGeneratePrePiledTree();
            
#if DEBUG
            ExportDebugFile(prePiledTree, "1_PrePiled.xml");
#endif
            // [Phase2] 스냅샷을 구성하기 전에 Core 및 선행 모드의 Defs를 미리 트리에 병합하여 
            // 뼈대(Schema)와 모드 Def 사이의 상속 체인을 이어줍니다.
            if (referenceDefsRoots != null && referenceDefsRoots.Count > 0)
            {
                Log.Msg("[Phase2] 참조 모드(Core 등) Defs 병합 중...");
                LoadAndMergeModDefs(prePiledTree, referenceDefsRoots, false);
            }
#if DEBUG
            ExportDebugFile(prePiledTree, "2_MergedWithReference.xml");
#endif
            // -------------------------------------------------------------------------
            // [Phase3] PostProcessor (Stage A 등)
            prePiledTree = XDocumentProcedureInjector.Instance.ExecuteStage(InjectionStage.StageA, prePiledTree);

            // -------------------------------------------------------------------------
            // [Phase4] LoadFolders.xml 조건(Any/All)에 따른 초기 스냅샷 멀티버스 생성
            Stopwatch stopwatch = Stopwatch.StartNew();
            Log.Msg("[Phase4] LoadFolders 조건(Any/All)에 따른 초기 스냅샷 분기 생성...");
            List<DefSnapshot> multiverse = DoctorStrange.GenerateInitialSnapshots(prePiledTree, selectedFolders);
            Log.Msg($"{stopwatch.ElapsedMilliseconds} ms 경과");
            // -------------------------------------------------------------------------
            // [Phase5] 각 스냅샷별로 자신에게 할당된 Defs 병합
            stopwatch.Restart();
            Log.Msg("[Phase5] 분기된 스냅샷별 Defs XML 로드 및 병합...");
            foreach (var snapshot in multiverse)
            {
                LoadAndMergeModDefs(snapshot.Tree, snapshot.AssignedFolders, true);
            }
            Log.Msg($"{stopwatch.ElapsedMilliseconds} ms 경과");
#if DEBUG
            ExportDebugFile(multiverse.First().Tree, "4_Merged.xml");
#endif

            // -------------------------------------------------------------------------
            // [Phase6] 각 스냅샷별 PatchOperation 분기 검사 및 추가 분열
            stopwatch.Restart();
            Log.Msg("[Phase6] PatchOperationFindMod 등 조건부 패치를 통한 스냅샷 2차 분열...");
            List<DefSnapshot> finalMultiverse = new List<DefSnapshot>();
            
            foreach (var snapshot in multiverse)
            {
                var branchedSnapshots = DoctorStrange.ProcessPatchOperations(snapshot, snapshot.AssignedPatches);
                finalMultiverse.AddRange(branchedSnapshots);
            }
            Log.Msg($"{stopwatch.ElapsedMilliseconds} ms 경과");
#if DEBUG
            ExportDebugFile(multiverse.First().Tree, "5_Patched.xml");
#endif

            // -------------------------------------------------------------------------
            // [Phase7] 상속(Inheritance) 처리 (모든 최종 평행 우주에 대해 각각 수행)
            // 이 단계에서 트리는 XML 생성 명세서에서 실제 트리 구조로 전환됩니다.
            stopwatch.Restart();
            Log.Msg("[Phase7] 최종 생성된 모든 스냅샷에 대해 XML 상속(ParentName) 처리...");
            foreach (var snapshot in finalMultiverse)
            {
                DoXmlInheritance(snapshot.Tree.Root);
                CleanUpAbstractNodes(snapshot.Tree.Root);
            }
            Log.Msg($"{stopwatch.ElapsedMilliseconds} ms 경과");
#if DEBUG
            ExportDebugFile(multiverse.First().Tree, "6_Inherited.xml");
#endif



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
        private static void LoadAndMergeModDefs(XDocument baseTree, IEnumerable<string> folderPaths, bool markAsTarget = false)
        {
            var baseRoot = baseTree.Root;
            if (baseRoot == null) return;

            foreach (var folderPath in folderPaths)
            {
                if (!Directory.Exists(folderPath)) continue;

                var xmlFiles = FileInterface.DescendantFiles(folderPath)
                    .Where(x => x.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
                
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var modDefDoc = FileInterface.ReadXml(xmlFile);
                        var modRoot = modDefDoc.Root;
                        if (modRoot != null && modRoot.Name.LocalName == "Defs")
                        {
                            var elements = modRoot.Elements().ToList();
                            
                            // 타겟 모드의 Def일 경우 추출 대상 마킹
                            if (markAsTarget)
                            {
                                foreach (var element in elements)
                                {
                                    element.SetAttributeValue("ExtractionTarget", "True");
                                }
                            }
                            
                            baseRoot.Add(elements);
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Wrn($"Def XML 병합 오류 ({xmlFile}): {e.Message}");
                    }
                }
            }
        }

        private static void LoadAndMergeModDefs(XDocument baseTree, List<ExtractableFolder> defFolders, bool markAsTarget = false)
        {
            LoadAndMergeModDefs(baseTree, defFolders.Select(f => f.FullPath), markAsTarget);
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

            // 1. 자기 자신이 뼈대 스키마(예: <JobDef Name="JobDef">)인지 검사하여 무한 루프 방지
            bool isBaseSchemaNode = def.Attribute("Name")?.Value == def.Name.LocalName;

            // 2. 부모가 없고 자기 자신이 뼈대가 아니라면, 자기 태그명으로 된 부모를 명시적으로 주입
            if (def.Attribute("ParentName") == null && !isBaseSchemaNode)
            {
                def.SetAttributeValue("ParentName", def.Name.LocalName);
            }

            // 3. 이제 모든 노드가 명시적인 ParentName을 가지게 되었으므로, 단일 로직으로 상속 처리
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

        private static void ExportDebugFile(XDocument singleTree, string fileNameWithoutExtension)
        {
            try
            {
                if (singleTree != null)
                {
                    string debugPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileNameWithoutExtension);
                    singleTree.Save(debugPath);
                    Log.Msg($"[디버그] 0번 스냅샷 트리를 출력했습니다: {debugPath}");
                }
            }
            catch (Exception e)
            {
                Log.Wrn($"[디버그] 스냅샷 트리 출력 중 오류 발생: {e.Message}");
            }
        }
    }
}