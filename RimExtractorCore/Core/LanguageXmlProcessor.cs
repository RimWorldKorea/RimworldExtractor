using System.Security;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore;

/// <summary>
/// 림월드 고유의 XML 구조와 관련된 처리를 위한 도구입니다.
/// </summary>
public static class LanguageXmlProcessor
{
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
                    if (!isOfficial && translation.FullListTranslate)
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
                        doc.Root!.AppendComment(
                            $"Original={SecurityElement.Escape(translation.Original).Replace('-', ' ')}");

                    var elem = doc.Root!.AppendElement(translation.Node,
                        translation.Translated ?? translation.Original);
                    ProcessPointerReplacement(elem, translations);
                }

                foreach (var (key, doc) in localXmls)
                {
                    var tokens = key.Split('|');
                    var className = tokens[0];
                    var outputPath = isOfficial
                        ? Path.Combine(targetFolder, className, tokens[1] + ".xml")
                        : Path.Combine(targetFolder, className,
                            Utils.GenerateFileName(Path.GetFileNameWithoutExtension(modName), className) + ".xml");
                    
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
                    doc.Root!.AppendComment(
                        $"Original={SecurityElement.Escape(translation.Original).Replace('-', ' ')}");

                var elem = doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
                ProcessPointerReplacement(elem, translations);
            }

            foreach (var (key, doc) in localXmls)
            {
                var tokens = key.Split('|');
                var className = tokens[0];
                var outputPath = isOfficial
                    ? Path.Combine(defInjectedDir, className, tokens[1] + ".xml")
                    : Path.Combine(defInjectedDir, className,
                        Utils.GenerateFileName(Path.GetFileNameWithoutExtension(modName), className) + ".xml");
                
                xmlFiles[outputPath] = doc;
            }
        }

        // 5. Full List Translations (XML)
        if (defInjectedFullListTranslations.Count > 0)
        {
    var defInjectedDir = Path.Combine(translationDir, "DefInjected");

            // 그룹화: ClassName과 부모 노드 이름(인덱스를 뺀 부분)을 기준으로 묶습니다.
            // 예: "RecipeDef.recipeUsers.0" -> "RecipeDef.recipeUsers"
            var groupedTranslations = defInjectedFullListTranslations
                .GroupBy(t =>
                {
                    int lastDot = t.Node.LastIndexOf('.');
                    string parentNode = lastDot > 0 ? t.Node.Substring(0, lastDot) : t.Node;
                    return new { t.ClassName, ParentNode = parentNode };
                });

            foreach (var group in groupedTranslations)
            {
                var className = group.Key.ClassName;
                var nodeParent = group.Key.ParentNode;

                var doc = new XDocument(new XElement("LanguageData"));
                
                // 부모 노드 생성 (예: <RecipeDef.recipeUsers>)
                var parentElem = new XElement(nodeParent);

                // 인덱스 번호순으로 정렬하여 <li> 요소 조립 (리스트 순서 보장)
                var sortedTranslations = group.OrderBy(t =>
                {
                    int lastDot = t.Node.LastIndexOf('.');
                    if (lastDot > 0 && int.TryParse(t.Node.Substring(lastDot + 1), out int index))
                        return index;
                    return 0;
                });

                foreach (var translation in sortedTranslations)
                {
                    if (commentOriginal)
                    {
                        // 주석을 <li> 태그 바로 위에 달아줍니다.
                        parentElem.AppendComment(
                            $"Original={SecurityElement.Escape(translation.Original).Replace('-', ' ')}");
                    }

                    // <li> 태그로 텍스트를 감싸서 생성
                    var liElem = new XElement("li", translation.Translated ?? translation.Original);
                    
                    // 포인터 교체 로직을 <li> 태그 자체에 적용
                    ProcessPointerReplacement(liElem, translations);
                    
                    // 부모 노드에 <li> 추가
                    parentElem.Add(liElem);
                }

                // 완성된 부모 노드를 루트(<LanguageData>)에 추가
                doc.Root!.Add(parentElem);

                // 파일 경로 생성 및 저장
                var outputPath = Path.Combine(defInjectedDir, className,
                    Utils.GenerateFileName(Path.GetFileNameWithoutExtension(modName), className, nodeParent) + ".xml");

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
                    doc.Root!.AppendComment(
                        $"{SettingManager.Current.OriginalLanguage}={SecurityElement.Escape(translation.Original).Replace('-', ' ')}");

                doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
            }

            foreach (var (key, doc) in localXmls)
            {
                var outputPath = isOfficial
                    ? Path.Combine(keyedDir, $"{key}.xml")
                    : Path.Combine(keyedDir,
                        Utils.GenerateFileName(Path.GetFileNameWithoutExtension(modName), "Keyed") + ".xml");
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

    /// <summary>
    /// Keyed 폴더의 XML을 읽어 TranslationEntry 리스트로 반환합니다.
    /// </summary>
    public static IEnumerable<TranslationEntry> ParseKeyed(string keyedDir, RequiredMods? requiredMods,
        bool isOfficialContent)
    {
        if (!Directory.Exists(keyedDir)) yield break;

        foreach (var xmlPath in FileInterface.DescendantFiles(keyedDir)
                     .Where(x => x.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            var fileName = isOfficialContent ? Path.GetFileNameWithoutExtension(xmlPath) : null;
            XDocument doc;
            try
            {
                doc = FileInterface.ReadXml(xmlPath);
            }
            catch
            {
                continue;
            }

            if (doc.Root == null) continue;

            foreach (var node in doc.Root.Elements())
            {
                yield return new TranslationEntry("Keyed", node.Name.LocalName, node.Value, null, requiredMods,
                    fileName);
            }
        }
    }

    /// <summary>
    /// Strings 폴더의 TXT를 읽어 TranslationEntry 리스트로 반환합니다.
    /// </summary>
    public static IEnumerable<TranslationEntry> ParseStrings(string stringsDir, RequiredMods? requiredMods)
    {
        if (!Directory.Exists(stringsDir)) yield break;

        foreach (var txtPath in FileInterface.DescendantFiles(stringsDir)
                     .Where(x => x.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)))
        {
            var nodeName = Path.GetRelativePath(stringsDir, txtPath);
            nodeName = Path.GetFileNameWithoutExtension(nodeName.Replace('\\', '.'));
            string[] lines;
            try
            {
                lines = File.ReadAllLines(txtPath);
            }
            catch
            {
                continue;
            }

            for (var i = 0; i < lines.Length; i++)
            {
                yield return new TranslationEntry("Strings", $"{nodeName}.{i}", lines[i], null, requiredMods, null);
            }
        }
    }
}