using System.Diagnostics;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.Procedures;

namespace RimExtractorCore.DefTreeSimulator;

/// <summary>
/// 림월드 엔진의 XML 로딩, 패치(PatchOperation), 상속(Inheritance) 과정을 모사하여
/// 추출기가 분석할 수 있는 최종 런타임 형태의 가상 DOM(DefSnapshot 리스트)을 생성합니다.
/// </summary>
public static class DefTreeSimulator
{
    public static Dictionary<string, XElement> DeepSchemaTypes { get; private set; } = new();

    // [리팩토링] 복잡했던 매개변수를 쳐내고, Mod 객체와 TargetVersion을 중심으로 개편
    public static SimulationResult Execute(Mod targetMod, Version targetVersion, List<Mod>? referenceMods = null)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        var result = new SimulationResult { TargetMod = targetMod };

        Log.Msg("[Phase1] 일차 스키마(PrePiledTree) 로드 및 생성...");
        // 1. 필요한 모든 DLL(어셈블리)을 긁어모아 스키마 트리 생성
        XDocument prePiledTree = LoadOrGeneratePrePiledTree(targetMod, referenceMods, targetVersion);

        // [Phase2] (생략 - Core Defs 통합 등 필요한 경우 추가 가능)
        // [Phase3] PostProcessor (Stage A)
        prePiledTree = DefTreeProcedureInjector.Instance.ExecuteStage(InjectionStage.StageA, prePiledTree);
        
        // [Phase4] 멀티버스(평행우주) 스냅샷 생성
        Log.Msg("[Phase4] LoadFolders 규칙 평가 및 초기 평행우주 생성...");
        var targetMods = new List<Mod> { targetMod };
        List<DefSnapshot> multiverse = DoctorStrange.GenerateInitialSnapshots(prePiledTree, targetMods, targetVersion);
        
        // [Phase5 & 6] 각 유니버스별 Defs 병합 및 PatchOperation 수행
        List<DefSnapshot> finalMultiverse = multiverse.AsParallel().Select(snapshot =>
        {
            // 5. 현재 우주에 로드된 폴더(LoadedUnits)들의 Defs 폴더 병합
            LoadAndMergeModDefs(snapshot.Tree, snapshot.LoadedUnits, true);
            
            // 6. 패치 연산 적용
            var patchedUniverses = DoctorStrange.ProcessPatchOperations(snapshot);
            return patchedUniverses;
        }).SelectMany(x => x).ToList();

        // [Phase7, 8 & 9] 후처리 병렬 수행
        Parallel.ForEach(finalMultiverse, snapshot =>
        {
            DoXmlInheritance(snapshot.Tree.Root); // 7. 상속 처리
            CleanUpAbstractNodes(snapshot.Tree.Root);
            InjectDeepSchemaRecursive(snapshot.Tree.Root); // 8. 딥 스키마 주입
            ApplyDefInjectedLanguageData(snapshot); // 9. 언어 데이터(DefInjected) 반영
        });

        Log.Msg($"{stopwatch.ElapsedMilliseconds/1000} ms 소요됨.");
        result.Snapshots = finalMultiverse;
        return result;
    }

    /// <summary>
    /// 어셈블리로부터 C# 클래스 구조를 미러링한 뼈대 트리를 가져옵니다.
    /// </summary>
    private static XDocument LoadOrGeneratePrePiledTree(Mod targetMod, List<Mod>? referenceMods, Version targetVersion)
    {
        var path = Engine.PrePiledTreePath;
        if (string.IsNullOrEmpty(path) || !File.Exists(path))
            throw new FileNotFoundException($"사전 컴파일된 XML을 찾을 수 없습니다: {path}");

        var doc = XDocument.Load(path);
        if (doc.Root != null && doc.Root.Name.LocalName == "PrePiled")
        {
            var typesNode = doc.Root.Element("Types");
            if (typesNode != null)
            {
                var assemblyPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                var allMods = new List<Mod> { targetMod };
                if (referenceMods != null) allMods.AddRange(referenceMods);

                foreach (var mod in allMods)
                {
                    if (string.IsNullOrWhiteSpace(mod.Path)) continue;
                    
                    // 대신 DoctorStrange의 순수 경로 평가기를 사용하여 문자열 경로만 받아옵니다.
                    var activePaths = DoctorStrange.EvaluateActivePaths(mod, targetVersion, new HashSet<string>());
                    
                    foreach (var relativePath in activePaths)
                    {
                        var dir = string.IsNullOrEmpty(relativePath) ? mod.Path : Path.Combine(mod.Path, relativePath);
                        var asmDir = Path.Combine(dir, "Assemblies");
                        
                        if (Directory.Exists(asmDir))
                        {
                            foreach (var dll in Directory.GetFiles(asmDir, "*.dll", SearchOption.AllDirectories))
                                assemblyPaths.Add(dll);
                        }
                    }
                }

                if (assemblyPaths.Count > 0)
                {
                    Log.Msg($"추가된 모드 어셈블리 {assemblyPaths.Count}개 스키마 병합 중...");
                    AssemblyResolver.AppendModAssembliesSchema(assemblyPaths, typesNode);
                }

                DoXmlInheritance(typesNode);
                DeepSchemaTypes.Clear();
                foreach (var typeElem in typesNode.Elements())
                    DeepSchemaTypes[typeElem.Name.LocalName] = typeElem;
            }
            var defsNode = doc.Root.Element("Defs");
            if (defsNode != null) return new XDocument(defsNode);
        }
        return doc;
    }

    /// <summary>
    /// 추출된 Defs 폴더들의 XML을 읽어 Base 트리에 병합합니다.
    /// </summary>
    private static void LoadAndMergeModDefs(XDocument baseTree, List<FolderLoadUnit> loadedUnits, bool markAsTarget = false)
    {
        var baseRoot = baseTree.Root;
        if (baseRoot == null) return;

        foreach (var unit in loadedUnits)
        {
            // [리팩토링] 하드코딩 패스 대신 FolderLoadUnit의 Defs 프로퍼티 직접 참조
            if (unit.Defs == null || !unit.Defs.Exists) continue;

            var xmlFiles = unit.Defs.GetFiles("*.xml", SearchOption.AllDirectories);
            foreach (var xmlFile in xmlFiles)
            {
                try
                {
                    var modDefDoc = FileInterface.ReadXml(xmlFile.FullName);
                    var modRoot = modDefDoc.Root;
                    if (modRoot != null && modRoot.Name.LocalName == "Defs")
                    {
                        var elements = modRoot.Elements().ToList();
                        if (markAsTarget)
                            foreach (var element in elements) element.SetAttributeValue("ExtractionTarget", "True");
                        baseRoot.Add(elements);
                    }
                }
                catch (Exception e)
                {
                    Log.Wrn($"Def XML 병합 실패 ({xmlFile.FullName}): {e.Message}");
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

    private static void ResolveInheritanceRecursive(XElement def, Dictionary<string, XElement> defsByName,
        HashSet<XElement> resolvedSet)
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
    /// 상속 처리가 모두 끝난 후, 불필요한 Abstract 노드를 삭제합니다.
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

        if (!string.IsNullOrEmpty(actualType))
        {
            // 2. 네임스페이스 분리 검사 (마지막 '.' 기준)
            int lastDotIndex = actualType.LastIndexOf('.');
            if (lastDotIndex > 0)
            {
                string namespacePart = actualType.Substring(0, lastDotIndex);
                
                // VIP 네임스페이스인 경우에만 앞부분을 날립니다. (모드 네임스페이스는 보존)
                if (Constants.IgnoredNamespaceNames.Contains(namespacePart))
                {
                    actualType = actualType.Substring(lastDotIndex + 1);
                }
            }

            // 3. XML 스키마 이름 규칙에 맞춰 중첩 클래스 기호(+) 치환 (예: ThingDef+Nested -> ThingDef.Nested)
            actualType = actualType.Replace('+', '.');
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
        
        foreach (var lang in priorityLanguages)
        {
            var englishName = lang.Split('(')[0].Trim();
            var languageInfo = LanguageInfo.Get(englishName);

            foreach (var unit in snapshot.LoadedUnits)
            {
                var defInjectedDir = unit.GetDefInjected(languageInfo); 
                if (defInjectedDir == null || !defInjectedDir.Exists) continue;

                foreach (var xmlFile in defInjectedDir.GetFiles("*.xml", SearchOption.AllDirectories))
                {
                    try
                    {
                        var className = Path.GetRelativePath(defInjectedDir.FullName, xmlFile.FullName)
                            .Split(Path.DirectorySeparatorChar).First();
                        var doc = FileInterface.ReadXml(xmlFile.FullName);
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
                    catch (Exception e)
                    {
                        Log.Wrn($"DefInjected 병합 실패 ({xmlFile.FullName}): {e.Message}");
                    }
                }
            }
        }
    }
    
    /// <summary>
    /// Defs 노드 하위에서 타겟 타입을 찾아 어트리뷰트를 주입합니다. (프로시저용)
    /// </summary>
    public static void GetDefAndSetAttribute(this XDocument doc, string targetType, string memberName, string attributeName, string value)
    {
        var defsNode = doc.Root?.Element("Defs");
        if (defsNode == null) return;
        
        // Utils에 추가한 GetOrCreateElement 활용
        var typeNode = defsNode.GetOrCreateElement(targetType);
        
        // 스키마 상속 구조 등을 위해 Name 속성이 없다면 기본으로 달아줍니다.
        if (typeNode.Attribute("Name") == null)
        {
            typeNode.SetAttributeValue("Name", targetType); 
        }

        var memberNode = typeNode.GetOrCreateElement(memberName);
        memberNode.SetAttributeValue(attributeName, value);
    }

    /// <summary>
    /// Types 노드 하위에서 타겟 타입을 찾아 어트리뷰트를 주입합니다. (프로시저용)
    /// </summary>
    public static void GetTypeAndSetAttribute(this XDocument doc, string targetType, string memberName, string attributeName, string value)
    {
        var typesNode = doc.Root?.Element("Types");
        if (typesNode == null) return;
        
        var typeNode = typesNode.GetOrCreateElement(targetType);
        //todo 여기에 Abstract도 추가해줘야 되는거 아닌가?
        
        if (typeNode.Attribute("Name") == null)
        {
            typeNode.SetAttributeValue("Name", targetType); 
        }

        var memberNode = typeNode.GetOrCreateElement(memberName);
        memberNode.SetAttributeValue(attributeName, value);
    }
}