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
    /// <summary>어셈블리에서 읽은 타입 정보를 저장해둡니다.</summary>
    public static Dictionary<string, XElement> DeepSchemaTypes { get; private set; } = new();
    
    public static SimulationResult Execute(Mod targetMod, Version targetVersion, List<Mod>? referenceMods = null)
    {
        var result = new SimulationResult { TargetMod = targetMod };

        // 1. 필요한 모든 DLL(어셈블리)을 긁어모아 스키마 트리 생성
        Log.Msg("사전 해석 트리 생성.", true);
        XDocument prePiledTree = LoadOrGeneratePrePiledTree(targetMod, referenceMods, targetVersion);
        
        // 분기 환경 스냅샷 생성
        Log.Msg("런타임 환경 분기 생성.", true);
        List<DefSnapshot> multiverse =
            DoctorStrange.GenerateInitialSnapshots(prePiledTree, targetMod, referenceMods, targetVersion);
        
        // [Phase5 & 6] 각 유니버스별 Defs 병합 및 PatchOperation 수행
        Log.Msg("Def 정의 및 Patch 병합.", true);
        var finalMultiverse = new List<DefSnapshot>();
        Log.RunWithCyclingDot(() =>
        {
            finalMultiverse = multiverse.AsParallel().Select(snapshot =>
            {
                // 5. 현재 우주에 로드된 폴더(LoadedUnits)들의 Defs 폴더 병합
                LoadAndMergeModDefs(snapshot.Tree, snapshot.LoadedUnits, targetMod);
                // 6. 패치 연산 적용
                var patchedUniverses = DoctorStrange.ProcessPatchOperations(snapshot);
                return patchedUniverses;
            }).SelectMany(x => x).ToList();
        });


        var _Debug1 = finalMultiverse.FirstOrDefault(x => x.IsBaseSnapshot);
        if (_Debug1 != null && _Debug1.Tree != null)
            ExportDebugFile(_Debug1.Tree, "Debug_BaseSnapshotTree_Inter2.xml");
        
        
        // [Phase7, 8 & 9] 후처리 병렬 수행
        Log.Msg("상속 처리 및 타입 정보 주입.", true);
        Log.RunWithCyclingDot(() =>
        {
            Parallel.ForEach(finalMultiverse, snapshot =>
            {
                DoXmlInheritance(snapshot.Tree.Root); // 7. 상속 처리
                if(snapshot.IsBaseSnapshot)
                    ExportDebugFile(snapshot.Tree, "Debug_BaseSnapshotTree_Inter3.xml");
                
                CleanUpAbstractNodes(snapshot.Tree.Root);
                // 여기까진 각 속성의 타입 정보는 주입되지 않습니다.
                InjectDeepSchemaRecursive(snapshot.Tree.Root); // 8. 딥 스키마 주입
                // 이제 각 속성이 어떤 타입인지 알 수 있습니다.
                ApplyDefInjectedLanguageData(snapshot); // 9. 언어 데이터(DefInjected) 반영
            
                // 프로시저 실행
                DefTreeProcedureInjector.Instance.ExecuteStage(InjectionStage.PostTreeGeneration, snapshot);
            });
        });
        
        var baseSnapshot = finalMultiverse.FirstOrDefault(x => x.IsBaseSnapshot);
        if (baseSnapshot != null && baseSnapshot.Tree != null)
        {
            ExportDebugFile(baseSnapshot.Tree, "Debug_BaseSnapshotTree_Inter5.xml");
        }
        
        result.Snapshots = finalMultiverse;
        return result;
    }

    /// <summary>어셈블리로부터 C# 클래스 구조를 미러링한 뼈대 트리를 가져옵니다.</summary>
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
                    
                    var activePaths =
                        DoctorStrange.EvaluateActivePaths(mod, targetVersion, new HashSet<string>());
                    
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
                    AssemblyResolver.AppendModAssembliesSchema(assemblyPaths, typesNode);

                // Types는 DeepSchemaTypes로 뺍니다.
                DoXmlInheritance(typesNode); // 이건 뭔 이유가 있어서 여기 필요했는데 기억이 안나
                DeepSchemaTypes.Clear();
                foreach (var typeElem in typesNode.Elements())
                    DeepSchemaTypes[typeElem.Name.LocalName] = typeElem;
            }
            
            // Defs는 반환값으로 뺍니다.
            var defsNode = doc.Root.Element("Defs");
            if (defsNode != null) return new XDocument(defsNode);
        }
        return doc;
    }

    /// <summary> 추출된 Defs 폴더들의 XML을 읽어 Base 트리에 병합합니다. </summary>
    private static void LoadAndMergeModDefs(XDocument baseTree, List<FolderLoadUnit> loadedUnits, Mod targetMod)
    {
        var baseRoot = baseTree.Root;
        if (baseRoot == null) return;
        
        var fileJobs = new List<(string FilePath, bool IsTarget)>();
        foreach (var unit in loadedUnits)
        {
            if (unit.Defs == null || !unit.Defs.Exists) continue;
        
            bool isTargetMod = unit.Parent.Equals(targetMod);
            var xmlFiles = unit.Defs.GetFiles("*.xml", SearchOption.AllDirectories);

            foreach (var file in xmlFiles)
                fileJobs.Add((file.FullName, isTargetMod));
        }

        var parsedResults = fileJobs.AsParallel().Select(job =>
            {
                try
                {
                    var modDefDoc = FileInterface.ReadXml(job.FilePath);
                    return (Doc: modDefDoc, job.IsTarget, job.FilePath);
                }
                catch (Exception)
                {
                    Log.Wrn($"{job.FilePath} 읽기 실패.", true);
                    return (Doc: (XDocument?)null, job.IsTarget, job.FilePath);
                }
            })
            .Where(result => result.Doc != null) // 실패한 파일 걸러내기
            .ToList();

        foreach (var result in parsedResults)
        {
            var docsRoot = result.Doc!.Root;
            if (docsRoot == null || docsRoot.Name.LocalName != "Defs") continue;
        
            var elements = docsRoot.Elements().ToList();
        
            if (result.IsTarget)
                foreach (var element in elements)
                    element.SetAttributeValue("ExtractionTarget", "True");
        
            baseRoot.Add(elements);
            // 병합이 끝난 후 더 이상 필요 없는 문서 껍데기의 참조를 끊어 GC가 수집할 수 있게 돕습니다.
            result.Doc.Root.RemoveNodes(); 
        }
    }
    
    /// <summary>XML 트리의 ParentName을 추적하여 상속 병합을 수행합니다.</summary>
    private static void DoXmlInheritance(XElement? root)
    {
        if (root == null) return;

        var allDefs = root.Elements().ToList();
        var resolvedSet = new HashSet<XElement>();

        try
        {
            // Name이 붙여져있다 => Property가 아니라 Type이다.
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

    /// <summary>Def 상속 관계를 재귀적으로 구성합니다.</summary>
    /// <param name="def"></param>
    /// <param name="defsByName"></param>
    /// <param name="resolvedSet"></param>
    private static void ResolveInheritanceRecursive(XElement def, Dictionary<string, XElement> defsByName,
        HashSet<XElement> resolvedSet)
    {
        if (resolvedSet.Contains(def)) return;

        // 1. 자기 자신이 뼈대 스키마(예: <JobDef Name="JobDef">)인지 검사하여 무한 루프 방지
        bool isBaseSchemaNode = def.Attribute("Name")?.Value == def.Name.LocalName;

        // 2. 부모가 없고 자기 자신이 뼈대가 아니라면, 자기 태그명으로 된 부모를 명시적으로 주입
        // e.g. <ThingDef .../> => <ThingDef ... ParentName="ThingDef" />
        if (def.Attribute("ParentName") == null && !isBaseSchemaNode)
            def.SetAttributeValue("ParentName", def.Name.LocalName);
        
        // 3. 이제 모든 노드가 명시적인 ParentName을 가지게 되었으므로, 단일 로직으로 상속 처리
        var parentNameAttr = def.Attribute("ParentName");
        if (parentNameAttr != null)
        {
            string parentName = parentNameAttr.Value;
            
            if (defsByName.TryGetValue(parentName, out var parentDef))
            {
                ResolveInheritanceRecursive(parentDef, defsByName, resolvedSet);

                // 부모는 타겟(추출 대상)이 아닌데 자식은 타겟인 경우 MayNotTranslate를 각인합니다.
                bool markMayNotTranslate = parentDef.Attribute("ExtractionTarget")?.Value != "True" &&
                                           def.Attribute("ExtractionTarget")?.Value == "True";

                CopyMissingElements(parentDef, def, markMayNotTranslate);
            }
        }

        resolvedSet.Add(def);
    }

    /// <summary>자식에게 없는 부모의 속성 자식에게 복사해줍니다.</summary>
    /// <param name="parent">속성을 넘겨줄 부모 노드입니다.</param>
    /// <param name="child">속성을 넘겨받을 자식 노드입니다.</param>
    /// <param name="markMayNotTranslate">true인 경우 MayNotTranslate 어트리뷰트를 각인합니다.</param>
    private static void CopyMissingElements(XElement parent, XElement child, bool markMayNotTranslate = false)
    {
        // 자식에게 없는 부모의 어트리뷰트를 복사해줍니다.
        foreach (var attr in parent.Attributes())
        {
            if (attr.Name == "Name" || attr.Name == "Abstract" || attr.Name == "ParentName") continue;
            if (child.Attribute(attr.Name) == null)
            {
                child.Add(new XAttribute(attr));
            }
        }

        // li 노드를 제외한 요소 복사
        foreach (var parentElem in parent.Elements().Where(e => e.Name.LocalName != "li"))
        {
            var childElem = child.Element(parentElem.Name);
            if (childElem == null)
            {
                var newElem = new XElement(parentElem);
                
                if (markMayNotTranslate)
                {
                    // 자기 자신에게 부여
                    newElem.SetAttributeValue(Constants.AttrMayNotTranslate, "True");

                    // 하위에 중첩된 모든 자손 노드(<graphicData> 안의 <texPath> 등)에게도 꼼꼼히 부여
                    foreach (var desc in newElem.Descendants())
                        desc.SetAttributeValue(Constants.AttrMayNotTranslate, "True");
                }
                child.Add(newElem);
            }
            else
            {
                // 재귀 호출 시에도 상태를 유지하여, 깊은 곳에 복사되는 요소도 태그를 부여받게 합니다.
                CopyMissingElements(parentElem, childElem, markMayNotTranslate);
            }
        }

        // 2. 리스트(li) 노드 병합 (림월드 바닐라 리스트 상속 모사)
        var parentListElements = parent.Elements("li").ToList();
        if (parentListElements.Any())
        {
            // 자식 컨테이너에 Inherit="False"가 명시되어 있다면 부모 리스트를 상속받지 않음
            bool inheritList = child.Attribute("Inherit")?.Value.ToLower() != "false";

            if (inheritList)
            {
                var liNodesToAdd = new List<XElement>();
                foreach (var parentLi in parentListElements)
                {
                    var newLi = new XElement(parentLi);
                    if (markMayNotTranslate)
                    {
                        newLi.SetAttributeValue(Constants.AttrMayNotTranslate, "True");
                        foreach (var desc in newLi.Descendants())
                        {
                            desc.SetAttributeValue(Constants.AttrMayNotTranslate, "True");
                        }
                    }
                    liNodesToAdd.Add(newLi);
                }

                // 자식의 첫 번째 li를 찾음
                var firstChildLi = child.Elements("li").FirstOrDefault();
                
                if (firstChildLi != null)
                {
                    // 자식의 리스트 요소들 '앞'에 부모 리스트 뭉치를 순서대로 삽입 (Prepend)
                    foreach (var newLi in liNodesToAdd)
                    {
                        firstChildLi.AddBeforeSelf(newLi);
                    }
                }
                else
                {
                    // 자식에게 li가 하나도 없다면 자식 컨테이너 내부에 통째로 추가
                    child.Add(liNodesToAdd);
                }
            }
        }
    }

    /// <summary>상속 처리가 모두 끝난 후, 불필요한 Abstract 노드를 삭제합니다.</summary>
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

    /// <summary>XDocument를 그대로 파일로 출력합니다.</summary>
    private static void ExportDebugFile(this XDocument singleTree, string fileNameWithoutExtension)
    {
        try
        {
            if (singleTree != null)
            {
                string debugPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileNameWithoutExtension);
                singleTree.Save(debugPath);
            }
        }
        catch (Exception e)
        {
            Log.Wrn(e.Message);
        }
    }

    /// <summary>미리 빼둔 어셈블리 해석 트리의 Types 정보를 주입합니다.</summary>
    /// <param name="node"></param>
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

    /// <summary>
    /// DefInjected 번역 데이터를 찾아 실제 XML Tree에 주입하고 Notice Tag를 제거합니다.
    /// </summary>
    private static void ApplyDefInjectedLanguageData(DefSnapshot snapshot)
    {
        // 주의: 우선순위가 높은 언어가 나중에 덮어써야 하므로 Reverse()를 적용합니다. (RimWorld의 동작 방식과 동일)
        var priorityLanguages = SettingManager.Current.GetLanguagePriorityList().Reverse().ToList();
        
        foreach (var lang in priorityLanguages)
        {
            var englishName = lang.Split('(')[0].Trim();
            if(!LanguageInfo.TryParse(englishName, out var languageInfo)) continue;

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
    
    /// <summary>Defs 노드 하위에서 타겟 타입을 찾아 어트리뷰트를 주입합니다. (프로시저용)</summary>
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

    /// <summary>Types 노드 하위에서 타겟 타입을 찾아 어트리뷰트를 주입합니다. (프로시저용)</summary>
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

    /// <summary>특정 DefType에 해당하는 모든 Def를 추출 대상에서 제외합니다.</summary>
    /// <param name="snapshot">작업 대상이 되는 Def 스냅샷</param>
    /// <param name="defTypeName">ThingDef, ModNamespace.CustomDef 형식의 DefType 이름입니다.</param>
    public static void BanDefType(this DefSnapshot snapshot, string defTypeName)
    {
        // 스냅샷의 XML 트리가 유효한지 검사 (<Defs> 루트 노드 확인)
        if (snapshot?.Tree?.Root == null) return;

        // 루트 노드 하위에서 태그명이 defTypeName과 정확히 일치하는 모든 요소를 가져옵니다.
        var targetDefs = snapshot.Tree.Root.Elements(defTypeName);

        foreach (var def in targetDefs)
        {
            // 추출 프로시저가 무시하도록 ExtractionTarget 속성을 "False"로 덮어씌우거나 새로 생성합니다.
            def.SetAttributeValue("ExtractionTarget", "False");
        }
    }
}