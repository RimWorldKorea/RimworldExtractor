using System.Security;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore
{
    public static class LanguageXmlProcessor
    {
        //TODO 리팩토링중 임시
        private static readonly HashSet<string> FullListTranslationTags = new() { "rulesFiles", "rulesStrings", "pathList" };
        
        /// <summary>
        /// 번역 데이터를 분석하여 저장해야 할 XML 파일들과 TXT 파일들의 내용을 경로와 함께 반환합니다.
        /// (물리적 디스크 I/O를 수행하지 않는 순수 데이터 가공 메서드)
        /// </summary>
        public static (Dictionary<string, XDocument> XmlFiles, Dictionary<string, List<string>> TxtFiles) 
            GenerateLanguageFiles(
                List<TranslationEntry> translations, 
                bool skipNoTranslation, 
                bool commentOriginal, 
                string modName, 
                string rootDirPath)
        {
            var xmlFiles = new Dictionary<string, XDocument>();
            var txtFiles = new Dictionary<string, List<string>>();

            var translationLang = SettingManager.Current.TranslationLanguage;
            var translationDir = Path.Combine(rootDirPath, "Languages", translationLang);

            // 1. 분류 버킷
            var conditionalDefInjected = new List<TranslationEntry>();
            var defInjected = new List<TranslationEntry>();
            var defInjectedFullListTranslations = new List<TranslationEntry>();
            var keyed = new List<TranslationEntry>();
            var strings = new List<TranslationEntry>();

            var isOfficial = translations.Any(x => x.SourceFile != null);
            if (isOfficial) Log.Msg("공식 콘텐츠 번역 파일 생성을 시작합니다.");

            // 2. 데이터 분류
            foreach (var rawTranslation in translations)
            {
                var translation = rawTranslation;
                var className = translation.ClassName;

                if (skipNoTranslation && className != "Strings" && string.IsNullOrEmpty(translation.Translated))
                    continue;

                if (className.StartsWith("Patches."))
                {
                    translation = translation with { ClassName = className.Substring("Patches.".Length) };
                    conditionalDefInjected.Add(translation);
                    continue;
                }

                switch (className)
                {
                    case "Keyed":
                        keyed.Add(translation);
                        break;
                    case "Strings":
                        strings.Add(translation);
                        break;
                    default:
                        if (!isOfficial && FullListTranslationTags.Any(translation.Node.Contains))
                            defInjectedFullListTranslations.Add(translation);
                        else
                            defInjected.Add(translation);
                        break;
                }
            }

            if (skipNoTranslation && conditionalDefInjected.Count == 0 && defInjected.Count == 0 &&
                keyed.Count == 0 && translations.Count > 0 && defInjectedFullListTranslations.Count == 0)
            {
                Log.Wrn("모든 번역이 비어있어(skipNoTranslation 활성화) 생성할 파일이 없습니다.");
            }

            // 3. Conditional DefInjected (XML)
            if (conditionalDefInjected.Count > 0)
            {
                var conditionalBaseDir = Path.Combine(rootDirPath, "ConditionalLanguages");
                var groupedByMods = conditionalDefInjected.GroupBy(x => x.RequiredMods?.ToString() ?? "Common");
                
                foreach (var group in groupedByMods)
                {
                    var packageIdFolder = group.Key.Replace("::", "_").Replace("/", "_").Replace('\\', '_');
                    var targetFolder = Path.Combine(conditionalBaseDir, packageIdFolder, translationLang, "DefInjected");
                    var localXmls = new Dictionary<string, XDocument>();

                    foreach (var translation in group)
                    {
                        var key = $"{translation.ClassName}|{translation.SourceFile}";
                        if (!localXmls.TryGetValue(key, out var doc))
                        {
                            doc = new XDocument(new XElement("LanguageData"));
                            localXmls[key] = doc;
                        }

                        if (commentOriginal)
                            doc.Root!.AppendComment($"Original={SecurityElement.Escape(translation.Original).Replace('-', ' ')}");

                        var elem = doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
                        ProcessPointerReplacement(elem, translations);
                    }

                    foreach (var (key, doc) in localXmls)
                    {
                        var tokens = key.Split('|');
                        var className = tokens[0];
                        var outputPath = isOfficial
                            ? Path.Combine(targetFolder, className, tokens[1] + ".xml")
                            : Path.Combine(targetFolder, className, Utils.GenerateFileName(Path.GetFileNameWithoutExtension(modName), className) + ".xml");

                        doc.DoFullListTranslation();
                        xmlFiles[outputPath] = doc;
                    }
                }
            }

            // 4. DefInjected (XML)
            if (defInjected.Count > 0)
            {
                var defInjectedDir = Path.Combine(translationDir, "DefInjected");
                var localXmls = new Dictionary<string, XDocument>();
                
                foreach (var translation in defInjected)
                {
                    var key = $"{translation.ClassName}|{translation.SourceFile}";
                    if (!localXmls.TryGetValue(key, out var doc))
                    {
                        doc = new XDocument(new XElement("LanguageData"));
                        localXmls[key] = doc;
                    }
                    if (commentOriginal)
                        doc.Root!.AppendComment($"Original={SecurityElement.Escape(translation.Original).Replace('-', ' ')}");

                    var elem = doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
                    ProcessPointerReplacement(elem, translations);
                }

                foreach (var (key, doc) in localXmls)
                {
                    var tokens = key.Split('|');
                    var className = tokens[0];
                    var outputPath = isOfficial
                        ? Path.Combine(defInjectedDir, className, tokens[1] + ".xml")
                        : Path.Combine(defInjectedDir, className, Utils.GenerateFileName(Path.GetFileNameWithoutExtension(modName), className) + ".xml");

                    doc.DoFullListTranslation();
                    xmlFiles[outputPath] = doc;
                }
            }

            // 5. Full List Translations (XML)
            if (defInjectedFullListTranslations.Count > 0)
            {
                var defInjectedDir = Path.Combine(translationDir, "DefInjected");
                var localXmls = new Dictionary<(string, string), XDocument>();
                
                foreach (var translation in defInjectedFullListTranslations)
                {
                    var nodeParent = translation.Node.Substring(0, translation.Node.LastIndexOf('.'));
                    var key = (translation.ClassName, nodeParent);
                    if (!localXmls.TryGetValue(key, out var doc))
                    {
                        doc = new XDocument(new XElement("LanguageData"));
                        localXmls[key] = doc;
                    }
                    if (commentOriginal)
                        doc.Root!.AppendComment($"Original={SecurityElement.Escape(translation.Original).Replace('-', ' ')}");

                    var elem = doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
                    ProcessPointerReplacement(elem, translations);
                }

                foreach (var (keyTuple, doc) in localXmls)
                {
                    var className = keyTuple.Item1;
                    var nodeParent = keyTuple.Item2;
                    var outputPath = Path.Combine(defInjectedDir, className,
                        Utils.GenerateFileName(Path.GetFileNameWithoutExtension(modName), className, nodeParent) + ".xml");

                    doc.DoFullListTranslation();
                    xmlFiles[outputPath] = doc;
                }
            }

            // 6. Keyed (XML)
            if (keyed.Count > 0)
            {
                var keyedDir = Path.Combine(translationDir, "Keyed");
                var localXmls = new Dictionary<string, XDocument>();
                
                foreach (var translation in keyed)
                {
                    var key = isOfficial ? translation.SourceFile! : "default";
                    if (!localXmls.TryGetValue(key, out var doc))
                    {
                        doc = new XDocument(new XElement("LanguageData"));
                        localXmls[key] = doc;
                    }
                    if (commentOriginal)
                        doc.Root!.AppendComment($"{SettingManager.Current.OriginalLanguage}={SecurityElement.Escape(translation.Original).Replace('-', ' ')}");

                    doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
                }

                foreach (var (key, doc) in localXmls)
                {
                    var outputPath = isOfficial
                        ? Path.Combine(keyedDir, $"{key}.xml")
                        : Path.Combine(keyedDir, Utils.GenerateFileName(Path.GetFileNameWithoutExtension(modName), "Keyed") + ".xml");
                    xmlFiles[outputPath] = doc;
                }
            }

            // 7. Strings (TXT)
            if (strings.Count > 0)
            {
                var stringDir = Path.Combine(translationDir, "Strings");
                var localTxts = new Dictionary<string, List<string>>();
                
                foreach (var translation in strings)
                {
                    var className = translation.Node.Substring(0, translation.Node.LastIndexOf('.'));
                    if (!localTxts.TryGetValue(className, out var lines))
                    {
                        lines = new List<string>();
                        localTxts[className] = lines;
                    }
                    lines.Add(translation.Translated ?? translation.Original);
                }

                foreach (var (className, lines) in localTxts)
                {
                    var key = className.Substring(0, className.LastIndexOf('.')).Replace('.', Path.DirectorySeparatorChar);
                    var outputPath = Path.Combine(stringDir, key);
                    var fileNameTxt = Path.Combine(outputPath, $"{className.Split('.').Last()}.txt");
                    txtFiles[fileNameTxt] = lines;
                }
            }

            return (xmlFiles, txtFiles);
        }

        private static void ProcessPointerReplacement(XElement elem, List<TranslationEntry> translations)
        {
            if (!elem.Value.Contains("{*")) return;
            elem.Value = Regex.Replace(elem.Value, "\\{\\*(.*?)\\}", match =>
            {
                var targetIdentifier = match.Groups[1].Value;
                var replacement = translations.FirstOrDefault(x => $"{x.ClassName}+{x.Node}" == targetIdentifier);
                if (replacement != null)
                    return replacement.Translated ?? replacement.Original;
                Log.Err($"Pointer: {targetIdentifier} 대상 식별자를 찾을 수 없습니다.");
                return "ERR";
            });
        }

        public static void DoFullListTranslation(this XDocument defInjectedDoc)
        {
            var patterns = FullListTranslationTags.Select(x => $".+?\\.{x}\\.\\d+").ToList();
            var fullListdic = new Dictionary<string, XElement>();
            var removedNodesDic = new Dictionary<string, List<XElement>>();

            foreach (var childNode in defInjectedDoc.Root!.Elements().ToList())
            {
                var nodeName = childNode.Name.LocalName;
                if (!patterns.Any(x => Regex.IsMatch(nodeName, x)))
                    continue;

                nodeName = nodeName.Substring(0, nodeName.LastIndexOf('.'));

                if (!fullListdic.TryGetValue(nodeName, out var fullList))
                {
                    fullList = new XElement(nodeName);
                    fullListdic[nodeName] = fullList;
                }
                if (!removedNodesDic.TryGetValue(nodeName, out var removedList))
                {
                    removedList = new List<XElement>();
                    removedNodesDic[nodeName] = removedList;
                }

                fullList.AppendElement("li", childNode.Value);
                removedList.Add(childNode);
            }

            foreach (var (key, fullListNode) in fullListdic)
            {
                var removedList = removedNodesDic[key];
                removedList.Last().AddAfterSelf(fullListNode);
                foreach (var xmlNode in removedList)
                {
                    xmlNode.Remove();
                }
            }
        }
        
        /// <summary>
        /// DefInjected XML 문서를 읽어 TranslationEntry 리스트로 파싱합니다. (순수 데이터 가공)
        /// </summary>
        public static IEnumerable<TranslationEntry> ParseDefInjected(XDocument doc, string className)
        {
            if (doc.Root == null) yield break;

            foreach (var node in doc.Root.Elements())
            {
                var name = node.Name.LocalName;
                if (node.Elements().Any())
                {
                    // <li> 태그 등 리스트 형태일 경우
                    var children = node.Elements().ToList();
                    for (int i = 0; i < children.Count; i++)
                    {
                        yield return new TranslationEntry(
                            className, 
                            $"{name}.{i}", 
                            string.Empty, 
                            children[i].Value, 
                            null, 
                            null);
                    }
                }
                else
                {
                    // 단일 노드일 경우
                    yield return new TranslationEntry(
                        className, 
                        name, 
                        string.Empty, 
                        node.Value, 
                        null, 
                        null);
                }
            }
        }
    }
}