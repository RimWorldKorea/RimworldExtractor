using System.Xml.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.Procedures;

namespace RimExtractorCore.DefTreeSimulator
{
    /// <summary>
    /// 림월드의 런타임 Def 계층을 시뮬레이션합니다.
    /// </summary>
    public static class SimulatorEngine
    {
        // 이 클래스에 멤버 변수 정의하지 마세요
        
        /// <summary>
        /// 림월드의 런타임 Def 계층을 시뮬레이션 모델을 반환합니다.
        /// </summary>
        /// <param name="modMetadata"></param>
        /// <param name="selectedFolders"></param>
        /// <param name="prePatches"></param>
        /// <param name="referenceDefsRoots"></param>
        /// <param name="isOfficialContent"></param>
        /// <param name="referenceMods"></param>
        /// <returns>이 단계에선 번역 같은거 고려하지 말고 시뮬레이션 모델을 그대로 반환시키세요</returns>
        public static SimulationResult Execute(
            ModMetadata modMetadata,
            List<ExtractableFolder> selectedFolders,
            List<ExtractableFolder> prePatches,
            List<string>? referenceDefsRoots,
            bool isOfficialContent,
            List<ModMetadata>? referenceMods = null)
        {
            var result = new SimulationResult();
            result.TargetMod = modMetadata;
            result.TargetFolders = selectedFolders;
            result.ReferenceMods = referenceMods ?? new List<ModMetadata>();

            // ---------------------------------------------------------------------------------------------------------
            // TODO [단계 0] 어셈블리로부터 사전 모델의 베이스 구성
            // ---------------------------------------------------------------------------------------------------------

            // 🟢 (A 지점) PostProcessors/PostA 스크립트 적용 (순수 XDocument 전달 및 반환)
            result.DefTree = XDocumentProcedureInjector.Instance.ExecuteStage(InjectionStage.StageA, result.DefTree);

            // ---------------------------------------------------------------------------------------------------------
            // [단계 1] 사전 참조 모델 구성
            // ---------------------------------------------------------------------------------------------------------
            if (referenceDefsRoots != null)
            {
                LoadReferenceDefs(result, referenceDefsRoots);
            }

            var defFolders = selectedFolders.Where(x => Path.GetFileName(x.FolderName) == "Defs").ToList();
            foreach (var extractableFolder in defFolders)
            {
                var defsRoot = extractableFolder.FullPath;
                var requiredPackageId = extractableFolder.RequiredPackageId;
                foreach (var filePath in FileInterface.DescendantFiles(defsRoot).Where(x => x.ToLower().EndsWith(".xml")))
                {
                    try
                    {
                        var fileName = Path.GetFileNameWithoutExtension(filePath);
                        var childDoc = FileInterface.ReadXml(filePath);
                        if (childDoc.Root == null) continue;

                        foreach (var node in childDoc.Root.Elements())
                        {
                            var newNode = new XElement(node);

                            if (requiredPackageId != null)
                            {
                                newNode.SetAttributeValue("RequiredPackageId", requiredPackageId);
                            }

                            if (isOfficialContent)
                            {
                                newNode.SetAttributeValue("SourceFile", fileName);
                            }
                            result.DefTree.Root!.Add(newNode);
                            var attributeName = node.Attribute("Name")?.Value;
                            if (attributeName != null)
                            {
                                result.ParentNodeLookUp[attributeName] = newNode;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Err($"{filePath}를 읽는 중 에러 발생, {e.Message}");
                    }
                }
            }

            // 🟢 (B 지점) PostProcessors/PostB 스크립트 적용
            result.DefTree = XDocumentProcedureInjector.Instance.ExecuteStage(InjectionStage.StageB, result.DefTree);

            // ---------------------------------------------------------------------------------------------------------
            // [단계 2] 패치 오퍼레이션 적용 (PrePatch 및 XML 상속)
            // ---------------------------------------------------------------------------------------------------------
            DoPrePatch(result, prePatches);
            DoXmlInheritance(result);

            // 🟢 (C 지점) PostProcessors/PostC 스크립트 적용
            result.DefTree = XDocumentProcedureInjector.Instance.ExecuteStage(InjectionStage.StageC, result.DefTree);

            // ---------------------------------------------------------------------------------------------------------
            // [단계 3] 랭귀지 데이터 오버라이드
            // ---------------------------------------------------------------------------------------------------------
            ApplyDefInjectedLanguages(result, ConfigManager.Current.GetLanguagePriorityList());

            // 🟢 (D 지점) PostProcessors/PostD 스크립트 적용
            result.DefTree = XDocumentProcedureInjector.Instance.ExecuteStage(InjectionStage.StageD, result.DefTree);

#if DEBUG
            result.DefTree.Save("DefTree_Debug.xml");
#endif
            return result;
        }

        private static void LoadReferenceDefs(SimulationResult simResult, List<string> referenceDefsRoots)
        {
            foreach (var referenceDefsRoot in referenceDefsRoots)
            {
                foreach (var filePath in FileInterface.DescendantFiles(referenceDefsRoot)
                             .Where(x => x.ToLower().EndsWith(".xml")))
                {
                    try
                    {
                        var childDoc = FileInterface.ReadXml(filePath);
                        if (childDoc.Root == null) continue;

                        foreach (var node in childDoc.Root.Elements())
                        {
                            var newNode = new XElement(node);
                            newNode.SetAttributeValue("Reference", "True");
                            simResult.DefTree.Root!.Add(newNode);
                            var attributeName = node.Attribute("Name")?.Value;
                            if (attributeName != null)
                            {
                                if (simResult.ParentNodeLookUp.ContainsKey(attributeName))
                                {
                                    Log.Wrn($"Parent 노드의 이름이 겹칩니다: {attributeName}. 나중 것으로 덮어씌웁니다.");
                                }
                                simResult.ParentNodeLookUp[attributeName] = newNode;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Err($"{filePath} 를 읽는 중 에러 발생, {e.Message}");
                        throw;
                    }
                }
            }
        }
        
        private static void DoPrePatch(SimulationResult simResult, List<ExtractableFolder> prePatches)
        {
            foreach (var patchDir in prePatches)
            {
                PatchOperations.ExecutePrePatches(simResult, patchDir);
            }

            foreach (var node in simResult.DefTree.Root!.Elements())
            {
                var attributeName = node.Attribute("Name")?.Value;
                if (attributeName != null)
                {
                    simResult.ParentNodeLookUp[attributeName] = node;
                }
            }
        }

        public static void DoXmlInheritance(SimulationResult simResult, IEnumerable<XElement>? customNodes = null)
        {
            var newDoc = new XDocument(new XElement("Defs"));
            var defs = newDoc.Root!;

            customNodes ??= simResult.DefTree.Root!.Elements().ToList();

            foreach (var node in customNodes)
            {
                if (node.Attribute("Abstract")?.Value.ToLower() == "true"
                    // || node.Attribute("Reference")?.Value.ToLower() == "true"
                    )
                    continue;

                var parentName = node.Attribute("ParentName")?.Value;
                if (parentName == null)
                {
                    defs.Add(new XElement(node));
                    continue;
                }
                var parentNodes = new Stack<XElement>();
                parentNodes.Push(node);
                while (true)
                {
                    if (parentName != null)
                    {
                        if (simResult.ParentNodeLookUp.TryGetValue(parentName, out var parentNode))
                        {
                            parentNodes.Push(parentNode);
                            parentName = parentNode.Attribute("ParentName")?.Value;
                        }
                        else
                        {
                            Log.Wrn($"자식 노드={node.Element("defName")?.Value ?? "UNKNOWN"}의 부모 노드={parentName}를 찾을 수 없었습니다. ");
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }

                var mergedNode = new XElement(node.Name);
                foreach (var attr in node.Attributes())
                {
                    mergedNode.SetAttributeValue(attr.Name, attr.Value);
                }
                mergedNode.SetAttributeValue("ParentName", null);

                while (parentNodes.Count > 0)
                {
                    var parentNode = parentNodes.Pop();
                    XmlOverwriteRecursive(mergedNode, parentNode);
                }

                var requiredPackageId = node.Attribute("RequiredPackageId")?.Value;
                if (requiredPackageId != null)
                {
                    mergedNode.SetAttributeValue("RequiredPackageId", requiredPackageId);
                }

                var sourceFile = node.Attribute("SourceFile")?.Value;
                if (sourceFile != null)
                {
                    mergedNode.SetAttributeValue("SourceFile", sourceFile);
                }

                if (node.Attribute("Reference")?.Value.ToLower() == "true")
                {
                    mergedNode.SetAttributeValue("Reference", "True");
                }

                defs.Add(mergedNode);
            }

            simResult.DefTree = newDoc;
        }

        private static void XmlOverwriteRecursive(XElement current, XElement other)
        {
            if (current.Name != other.Name)
            {
                // Log.Wrn($"Different name, Original={Original.Name}|{Original.OuterXml}, other={other.Name}|{other.OuterXml}");
                return;
            }

            foreach (var otherChildNode in other.Elements())
            {
                var existingChildNode = current.Element(otherChildNode.Name);
                // 1. 존재하지 않을 경우
                if (existingChildNode == null)
                {
                    current.Add(new XElement(otherChildNode));
                    continue;
                }

                // 2. 상속을 원하지 않을 경우
                var inherit = otherChildNode.Attribute("Inherit")?.Value.ToLower() != "false";
                if (!inherit)
                {
                    existingChildNode.Remove();
                    current.Add(new XElement(otherChildNode));
                    continue;
                }

                // 3. 텍스트 노드 하나일 경우
                if (existingChildNode.IsTextNode())
                {
                    existingChildNode.Remove();
                    current.Add(new XElement(otherChildNode));
                    continue;
                }
                // 4. 리스트 노드일 경우 (상속 대상이나 기존 노드가 리스트 노드 형태인 경우)
                if (otherChildNode.Elements().Any(x => x.IsListNode()) || existingChildNode.Elements().Any(x => x.IsListNode()))
                {
                    foreach (var childNode in otherChildNode.Elements())
                    {
                        existingChildNode.Add(new XElement(childNode));
                    }

                    continue;
                }
                XmlOverwriteRecursive(existingChildNode, otherChildNode);
            }
        }

        /// <summary>
        /// 🟢 ModMetadata 및 ExtractableFolder 목록으로부터 LoadFolders 구조가 반영된 Languages/DefInjected 디렉터리들을 탐색합니다.
        /// </summary>
        private static List<string> GetCandidateDefInjectedDirs(ModMetadata mod, IEnumerable<ExtractableFolder>? folders, string lang)
        {
            var results = new List<string>();
            if (string.IsNullOrWhiteSpace(lang)) return results;

            var langShort = lang.Split(' ').First();
            var baseDirs = new HashSet<string>();

            // 1. ExtractableFolder의 FullPath 상위 경로 (LoadFolder 위치) 수집
            if (folders != null)
            {
                foreach (var folder in folders)
                {
                    var parentDir = Path.GetDirectoryName(folder.FullPath);
                    if (!string.IsNullOrEmpty(parentDir))
                    {
                        baseDirs.Add(parentDir);
                    }
                }
            }

            // 2. Fallback: 모드 루트 경로 추가
            baseDirs.Add(mod.RootDir);

            // 3. 각 LoadFolder 기반 디렉터리의 Languages/{lang}/DefInjected 존재 여부 검사
            foreach (var baseDir in baseDirs)
            {
                var pathFull = Path.Combine(baseDir, "Languages", lang, "DefInjected");
                if (Directory.Exists(pathFull)) results.Add(pathFull);

                var pathShort = Path.Combine(baseDir, "Languages", langShort, "DefInjected");
                if (Directory.Exists(pathShort)) results.Add(pathShort);
            }

            return results.Distinct().ToList();
        }

        /// <summary>
        /// 🟢 우선순위 언어 목록에 따라 simResult에 저장된 모드들의 DefInjected XML을 순회하며 DefTree를 덮어씁니다.
        /// (DefType별 defName 중복 및 LoadFolders.xml 하위 경로를 지원하도록 수정한 버전)
        /// </summary>
        public static void ApplyDefInjectedLanguages(
            SimulationResult simResult,
            IEnumerable<string> languages)
        {
            Log.Msg("Flag1");
            
            if (simResult.DefTree.Root == null) return;

            // 🟢 탐색할 (ModMetadata, Folders) 쌍 수집
            var modFolderPairs = new List<(ModMetadata Mod, IEnumerable<ExtractableFolder>? Folders)>();

            if (simResult.ReferenceMods != null)
            {
                foreach (var refMod in simResult.ReferenceMods)
                {
                    var refFolders = ModLister.GetExtractableFolders(refMod).Where(x => x.IsAutoSelectable());
                    modFolderPairs.Add((refMod, refFolders));
                }
            }

            if (simResult.TargetMod != null)
            {
                modFolderPairs.Add((simResult.TargetMod, simResult.TargetFolders));
            }

            if (modFolderPairs.Count == 0) return;

            // 🟢 defName을 키로 하고, (DefType, Element) 튜플 리스트를 값으로 갖는 LookUp 맵 생성
            var defLookup = new Dictionary<string, List<(string DefType, XElement Element)>>();
            foreach (var defNode in simResult.DefTree.Root.Elements())
            {
                var defName = defNode.Element("defName")?.Value;
                if (!string.IsNullOrEmpty(defName))
                {
                    var defType = defNode.Name.LocalName; // 예: "ThingDef", "JobDef", "StatDef"
                    if (!defLookup.TryGetValue(defName, out var list))
                    {
                        list = new List<(string DefType, XElement Element)>();
                        defLookup[defName] = list;
                    }

                    list.Add((defType, defNode));
                }
            }

            bool isFirstLanguage = true;

            foreach (var lang in languages)
            {
                if (string.IsNullOrWhiteSpace(lang)) continue;

                // 1차 언어는 완전히 덮어쓰고(overwrite = true), 2차 언어부터는 빈 곳만 채움(overwrite = false)
                bool overwrite = isFirstLanguage;
                isFirstLanguage = false;

                foreach (var (mod, folders) in modFolderPairs)
                {
                    Log.Msg($"Flag3:{mod.ModName}");
                    var candidateDirs = GetCandidateDefInjectedDirs(mod, folders, lang);

                    foreach (var dir in candidateDirs)
                    {
                        if (!Directory.Exists(dir))
                        {
#if DEBUG
                            Log.Msg($"\"{dir}\" 경로가 존재하지 않습니다.");
#endif
                            continue;
                        }

                        foreach (var filePath in FileInterface.DescendantFiles(dir).Where(x => x.ToLower().EndsWith(".xml")))
                        {
                            try
                            {
                                var doc = FileInterface.ReadXml(filePath);
                                if (doc.Root is null) continue;

                                // 🟢 상대 경로의 첫 번째 폴더명을 통해 DefType 추출
                                // 예: DefInjected/ThingDef/3vdl0p.xml -> "ThingDef"
                                var relPath = Path.GetRelativePath(dir, filePath);
                                var pathTokens =
                                    relPath.Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                                        StringSplitOptions.RemoveEmptyEntries);
                                var fileDefType = pathTokens.Length > 1 ? pathTokens[0] : null;

                                foreach (var node in doc.Root.Elements())
                                {
                                    var fullKey = node.Name.LocalName; // 예: "ULS_FlickLiftStructure.label"
                                    var textValue = node.Value;
                                    if (string.IsNullOrWhiteSpace(textValue)) continue;

                                    var dotIdx = fullKey.IndexOf('.');
                                    if (dotIdx <= 0) continue;

                                    var defName = fullKey[..dotIdx];
                                    var subPath = fullKey[(dotIdx + 1)..];
                                    
                                    Log.Msg($"defName: {defName} || propertyName: {subPath}");

                                    if (defLookup.TryGetValue(defName, out var matchingDefs))
                                    {
                                        // 🟢 1. 파일의 DefType(폴더명)과 일치하는 DefTarget 필터링
                                        var targetDefs = matchingDefs
                                            .Where(x => fileDefType != null && string.Equals(x.DefType, fileDefType,
                                                StringComparison.OrdinalIgnoreCase))
                                            .Select(x => x.Element)
                                            .ToList();

                                        // 🟢 2. 폴더 구조가 예외적인 경우(DefInjected 직하위에 xml이 있는 등) fallback 처리
                                        if (targetDefs.Count == 0)
                                        {
                                            targetDefs = matchingDefs.Select(x => x.Element).ToList();
                                        }

                                        foreach (var targetDef in targetDefs)
                                        {
                                            ApplyPathValue(targetDef, subPath.Split('.'), textValue, overwrite);
                                        }
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                Log.Err($"DefInjected 언어팩 읽기 실패 ({filePath}): {e.Message}");
                            }
                        }
                    }
                }
            }
        }

        private static void ApplyPathValue(XElement targetDef, string[] pathTokens, string value, bool overwrite)
        {
            XElement current = targetDef;
            for (int i = 0; i < pathTokens.Length; i++)
            {
                var token = pathTokens[i];
                bool isLast = (i == pathTokens.Length - 1);

                if (isLast)
                {
                    var elem = FindChildElement(current, token);
                    if (elem == null)
                    {
                        current.Add(new XElement(token, value));
                    }
                    else if (overwrite || string.IsNullOrWhiteSpace(elem.Value))
                    {
                        elem.Value = value;
                    }
                }
                else
                {
                    var nextElem = FindChildElement(current, token);
                    if (nextElem == null)
                    {
                        if (int.TryParse(token, out var idx))
                        {
                            var listElements = current.Elements("li").ToList();
                            while (listElements.Count <= idx)
                            {
                                var newLi = new XElement("li");
                                current.Add(newLi);
                                listElements.Add(newLi);
                            }
                            nextElem = listElements[idx];
                        }
                        else
                        {
                            nextElem = new XElement(token);
                            current.Add(nextElem);
                        }
                    }
                    current = nextElem;
                }
            }
        }
        
        private static XElement? FindChildElement(XElement parent, string token)
        {
            if (int.TryParse(token, out var idx))
            {
                var listElements = parent.Elements("li").ToList();
                if (idx >= 0 && idx < listElements.Count)
                    return listElements[idx];
                return null;
            }

            var child = parent.Element(token);
            if (child != null) return child;

            foreach (var handleChild in parent.Elements())
            {
                if (handleChild.HasElements)
                {
                    foreach (var sub in handleChild.Elements())
                    {
                        if (sub.IsTextNode() && sub.Value.EndsWith(token, StringComparison.OrdinalIgnoreCase))
                        {
                            return handleChild;
                        }
                    }
                }
            }

            return null;
        }
    }
}