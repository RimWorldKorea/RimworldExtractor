using System.Diagnostics;
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
        public static Dictionary<string, XElement> DeepSchemaTypes { get; private set; } = new();
        
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
            XDocument prePiledTree = LoadOrGeneratePrePiledTree(modMetadata, referenceMods, selectedFolders);
            
#if DEBUG
            ExportDebugFile(prePiledTree, "1_PrePiled.xml");
            
            var typesDebugTree = new XDocument(new XElement("Types", DeepSchemaTypes.Values));
            ExportDebugFile(typesDebugTree, "1_PrePiled_Types.xml");
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
                Log.Common($"페이즈5 {String.Join(",",snapshot.RequiredModIds)}");
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
                Log.Common($"페이즈6 {String.Join(",",snapshot.RequiredModIds)}");
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
                Log.Common($"페이즈7_1 {String.Join(",",snapshot.RequiredModIds)}");
                DoXmlInheritance(snapshot.Tree.Root);
                Log.Common($"페이즈7_2 {String.Join(",",snapshot.RequiredModIds)}");
                CleanUpAbstractNodes(snapshot.Tree.Root);
                Log.Common($"페이즈7_3 {String.Join(",",snapshot.RequiredModIds)}");
            }
            Log.Msg($"{stopwatch.ElapsedMilliseconds} ms 경과");
            
            // -------------------------------------------------------------------------
            // [Phase8] 딥 스키마(Deep Schema) 전파
            stopwatch.Restart();
            Log.Msg("[Phase8] 복합 타입 딥 스키마(Deep Schema) 전파 중...");
            foreach (var snapshot in finalMultiverse)
            {
                InjectDeepSchemaRecursive(snapshot.Tree.Root);
                Log.Common($"페이즈8 {String.Join(",",snapshot.RequiredModIds)}");
            }
            Log.Msg($"{stopwatch.ElapsedMilliseconds} ms  ");
            
#if DEBUG
            ExportDebugFile(multiverse.First().Tree, "6_Inherited.xml");
#endif
// -------------------------------------------------------------------------
            // [NEW] [Phase9] DefInjected LanguageData 병합
            stopwatch.Restart();
            Log.Msg("[Phase9] DefInjected 언어 데이터(LanguageData) 병합...");
            foreach (var snapshot in finalMultiverse)
            {
                ApplyDefInjectedLanguageData(snapshot);
            }
            Log.Msg($"{stopwatch.ElapsedMilliseconds} ms 소요됨.");


            result.Snapshots = finalMultiverse;
            return result;
        }

        /// <summary>
        /// 어셈블리로부터 C# 클래스 구조를 미러링한 뼈대 트리를 가져옵니다.
        /// </summary>
        private static XDocument LoadOrGeneratePrePiledTree(ModMetadata targetMod, List<ModMetadata>? referenceMods, List<ExtractableFolder> selectedFolders)
        {
            var path = ExtractorCore.PrePiledTreePath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                throw new FileNotFoundException($"통합 뼈대 XML 파일을 찾을 수 없습니다: {path}");
            }

            var doc = XDocument.Load(path);

            if (doc.Root != null && doc.Root.Name.LocalName == "PrePiled")
            {
                var typesNode = doc.Root.Element("Types");
                if (typesNode != null)
                {
                    // [수정됨] 탐색할 모든 "진짜 로드 폴더 루트"를 담을 집합 (중복 방지)
                    var searchDirectories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    // 1. [타겟 모드] 루트 폴더와 선택된 폴더들의 진짜 뿌리 추가
                    searchDirectories.Add(targetMod.RootDir);
                    foreach (var folder in selectedFolders)
                    {
                        searchDirectories.Add(folder.ActualLoadFolderRoot);
                    }

                    // 2. [참조 모드] 루트 폴더와 ModLister가 찾아낸 진짜 뿌리들 추가
                    if (referenceMods != null)
                    {
                        foreach (var refMod in referenceMods)
                        {
                            searchDirectories.Add(refMod.RootDir);
                            
                            foreach (var rf in ModLister.GetExtractableFolders(refMod))
                            {
                                searchDirectories.Add(rf.ActualLoadFolderRoot);
                            }
                        }
                    }

                    // 3. 수집된 모든 진짜 루트 경로에서 Assemblies 폴더를 견고하게 탐색
                    var assemblyPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    
                    Log.Msg($"{searchDirectories.Count}개 잠재적 어셈블리 보유 폴더(루트 및 버전 폴더) 탐색 중...");
                    foreach (var baseDir in searchDirectories)
                    {
                        var asmDir = Path.Combine(baseDir, "Assemblies");
                        // Log.Msg($"{asmDir} 보는 중"); // 필요 시 주석 해제
                        if (Directory.Exists(asmDir))
                        {
                            Log.Msg($"[DLL 탐색] 어셈블리 폴더 발견: {asmDir}");
                            foreach (var dll in Directory.GetFiles(asmDir, "*.dll", SearchOption.AllDirectories))
                            {
                                assemblyPaths.Add(dll);
                            }
                        }
                    }

                    // [NEW] 수집된 모드 DLL들을 메모리의 typesNode에 병합합니다!
                    if (assemblyPaths.Count > 0)
                    {
                        Log.Msg($"총 {assemblyPaths.Count}개의 어셈블리를 딥 스키마 딕셔너리에 병합합니다...");
                        AssemblyResolver.AppendModAssembliesSchema(assemblyPaths, typesNode);
                    }

                    // 코어 + 모드 C# 클래스 전체를 대상으로 상속(ParentName) 처리 시작!
                    DoXmlInheritance(typesNode);

                    DeepSchemaTypes.Clear();
                    foreach (var typeElem in typesNode.Elements())
                    {
                        // GroupBy-Last 로직과 동일하게 딕셔너리 덮어쓰기 할당
                        DeepSchemaTypes[typeElem.Name.LocalName] = typeElem;
                    }
                }

                var defsNode = doc.Root.Element("Defs");
                if (defsNode != null)
                {
                    return new XDocument(defsNode); 
                }
            }

            return doc;
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

            try
            {
                var defsByName = allDefs
                    .Where(e => e.Attribute("Name") != null)
                    .ToDictionary(e => e.Attribute("Name")!.Value, e => e);

                foreach (var def in allDefs)
                {
                    ResolveInheritanceRecursive(def, defsByName, resolvedSet);
                }
            }
            catch (Exception e)
            {
                Log.Err(e.ToString());
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
                    
                    // [NEW] 핵심 로직: 부모는 타겟(추출 대상)이 아닌데 자식은 타겟인 경우 true가 됩니다.
                    bool markMayNotTranslate = parentDef.Attribute("ExtractionTarget")?.Value != "True" && 
                                               def.Attribute("ExtractionTarget")?.Value == "True";
                    
                    CopyMissingElements(parentDef, def, markMayNotTranslate);
                }
            }
            
            resolvedSet.Add(def);
        }

        private static void CopyMissingElements(XElement parent, XElement child, bool markMayNotTranslate = false)
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
                    var newElem = new XElement(parentElem);
                    
                    // [NEW] 외부(비대상)에서 복사되어 들어오는 요소에 Notice Tag 부여
                    if (markMayNotTranslate)
                    {
                        // 자기 자신에게 부여
                        newElem.SetAttributeValue(Constants.AttrMayNotTranslate, "True");
                        
                        // 하위에 중첩된 모든 자손 노드(<graphicData> 안의 <texPath> 등)에게도 꼼꼼히 부여
                        foreach (var desc in newElem.Descendants())
                        {
                            desc.SetAttributeValue(Constants.AttrMayNotTranslate, "True");
                        }
                    }
                    
                    child.Add(newElem);
                }
                else
                {
                    // 재귀 호출 시에도 상태를 유지하여, 깊은 곳에 복사되는 요소도 태그를 부여받게 합니다.
                    CopyMissingElements(parentElem, childElem, markMayNotTranslate);
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
        
        private static void InjectDeepSchemaRecursive(XElement? node)
        {
            if (node == null) return;

            // 1. 현재 노드의 '진짜 타입(Actual Type)'을 알아냅니다.
            string? actualType = node.Attribute("Class")?.Value; // 다형성(Class) 최우선 확인
            
            // 모드의 커스텀 타입처럼 네임스페이스가 섞여있다면 뒤의 클래스명만 추출 
            // (예: Universal_Lift_Structure.CompProperties_LiftConsole -> CompProperties_LiftConsole)
            if (!string.IsNullOrEmpty(actualType) && actualType.Contains('.'))
            {
                actualType = actualType.Split('.').Last();
            }

            // Class가 없다면 Type 어트리뷰트 확인
            if (string.IsNullOrEmpty(actualType))
            {
                actualType = node.Attribute("Type")?.Value;
            }

            // li 노드인데 Class나 Type이 없다면 부모의 Type을 상속 
            // (예: <comps List="True" Type="CompProperties"> 의 자식 <li>)
            if (string.IsNullOrEmpty(actualType) && node.Name.LocalName == "li" && 
                node.Parent != null && node.Parent.Attribute("List")?.Value == "True")
            {
                actualType = node.Parent.Attribute("Type")?.Value;
            }

            // 2. 알아낸 타입에 해당하는 스키마가 캐시에 존재한다면 속성 전파
            if (!string.IsNullOrEmpty(actualType) && DeepSchemaTypes.TryGetValue(actualType, out var schemaNode))
            {
                // 자식 노드들에게 스키마의 속성을 물려줍니다.
                foreach (var targetChild in node.Elements())
                {
                    var childSchema = schemaNode.Element(targetChild.Name.LocalName);
                    if (childSchema != null)
                    {
                        // 스키마에 정의된 어트리뷰트(Type, NoTranslate, List 등) 복사
                        foreach (var attr in childSchema.Attributes())
                        {
                            if (targetChild.Attribute(attr.Name) == null)
                            {
                                targetChild.SetAttributeValue(attr.Name, attr.Value);
                            }
                        }
                    }
                }
            }

            // 3. 자식들로 계속 타고 들어갑니다 (재귀)
            foreach (var child in node.Elements())
            {
                InjectDeepSchemaRecursive(child);
            }
        }
        
        // [NEW] DefInjected 번역 데이터를 찾아 실제 XML Tree에 주입하고 Notice Tag를 제거하는 메서드
        private static void ApplyDefInjectedLanguageData(DefSnapshot snapshot)
        {
            // 주의: 우선순위가 높은 언어가 나중에 덮어써야 하므로 Reverse()를 적용합니다. (RimWorld의 동작 방식과 동일)
            var priorityLanguages = SettingManager.Current.GetLanguagePriorityList().Reverse().ToList();
            
            // [수정됨] 꼬리 자르기 역산 없이 객체의 프로퍼티를 즉시 사용!
            var versionDirs = snapshot.AssignedFolders
                .SelectMany(f => new[] { f.ActualLoadFolderRoot, f.Root.RootDir })
                .Where(d => !string.IsNullOrEmpty(d))
                .Distinct()
                .ToList();

            foreach (var lang in priorityLanguages)
            {
                var shortLang = lang.Split(' ').First();
                var langNames = new HashSet<string> { lang, shortLang };

                foreach (var versionDir in versionDirs)
                {
                    foreach (var langName in langNames)
                    {
                        var defInjectedDir = Path.Combine(versionDir, "Languages", langName, "DefInjected");
                        if (!Directory.Exists(defInjectedDir)) continue;

                        foreach (var xmlPath in FileInterface.DescendantFiles(defInjectedDir).Where(x => x.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
                        {
                            try
                            {
                                var className = Path.GetRelativePath(defInjectedDir, xmlPath).Split(Path.DirectorySeparatorChar).First();
                                var doc = FileInterface.ReadXml(xmlPath);
                                var parsed = LanguageXmlProcessor.ParseDefInjected(doc, className);
                                
                                foreach (var p in parsed)
                                {
                                    // Utils.GetXpath를 이용해 XML 내부의 실제 타겟 노드 경로를 가져옵니다.
                                    var xpath = Utils.GetXpath(p.ClassName, p.Node);
                                    var targetNodes = snapshot.Tree.SelectNodesSafe(xpath);
                                    
                                    if (targetNodes != null)
                                    {
                                        foreach (var targetNode in targetNodes)
                                        {
                                            var text = p.Translated ?? p.Original;
                                            if (!string.IsNullOrEmpty(text))
                                            {
                                                targetNode.Value = text; // 모드 언어 파일의 텍스트로 오버라이드
                                            }
                                            
                                            // 핵심: LanguageData에서 명시적으로 정의된 텍스트이므로, 상속 과정에서 붙었던 태그를 제거합니다!
                                            targetNode.Attribute(Constants.AttrMayNotTranslate)?.Remove();
                                        }
                                    }
                                }
                            }
                            catch (Exception e) { Log.Wrn($"DefInjected 병합 실패 ({xmlPath}): {e.Message}"); }
                        }
                    }
                }
            }
        }
    }
}