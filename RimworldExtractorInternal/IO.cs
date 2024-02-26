﻿using ClosedXML.Excel;
using RimworldExtractorInternal.Records;
using System.Security;
using System.Text.RegularExpressions;
using System.Xml;

namespace RimworldExtractorInternal
{
    public static class IO
    {
        private static readonly string HeaderClassNode = "Class+Node [(Identifier (Key)]";
        private static readonly string HeaderClass = "Class [Not chosen]";
        private static readonly string HeaderNode = "Node [Not chosen]";
        private static readonly string HeaderRequiredMods = "Required Mods [Not chosen]";
        private static string HeaderOriginal => $"{Prefabs.OriginalLanguage} [Source string]";
        private static string HeaderTranslated => $"{Prefabs.TranslationLanguage} [Translation]";
        public static void ToExcel(List<TranslationEntry> translations, string outPath = "result")
        {
            var xlsx = new XLWorkbook();
            var sheet = xlsx.AddWorksheet();
            sheet.Cell(1, 1).Value = HeaderClassNode;
            sheet.Cell(1, 2).Value = HeaderClass;
            sheet.Cell(1, 3).Value = HeaderNode;
            sheet.Cell(1, 4).Value = HeaderRequiredMods;
            sheet.Cell(1, 5).Value = HeaderOriginal;
            sheet.Cell(1, 6).Value = HeaderTranslated;
            for (int i = 0; i < translations.Count; i++)
            {
                var entry = translations[i];
                sheet.Cell(2 + i, 1).Value = $"{entry.className}+{entry.node}";
                sheet.Cell(2 + i, 2).Value = entry.className;
                sheet.Cell(2 + i, 3).Value = entry.node;
                if (entry.requiredMods != null)
                {
                    var combinedRequiredMods = string.Join('\n', entry.requiredMods);
                    sheet.Cell(2 + i, 4).Value = combinedRequiredMods;
                    if (combinedRequiredMods.Contains("##packageId##") && entry.className.StartsWith("Patches"))
                    {
                        Log.WrnOnce($"Required Mods 열에 잘못된 값이 존재합니다. 추후 Patches의 올바른 생성을 위해 엑셀 파일에 있는 해당 문구: \"{combinedRequiredMods}\" 를 직접 모드 이름으로 바꿔야 합니다.",
                            $"잘못된{combinedRequiredMods}경고".GetHashCode());
                    }
                }
                sheet.Cell(2 + i, 5).Value = entry.original;
            }

            sheet.Style.Font.FontName = "맑은 고딕";
            xlsx.SaveSafely(outPath + ".xlsx");
        }

        public static List<TranslationEntry> FromExcel(string inputPath)
        {
            var xlsx = new XLWorkbook(inputPath);
            var sheet = xlsx.Worksheets.Worksheet(1);
            var translations = new List<TranslationEntry>();
            var rows = sheet.RowsUsed().ToList();
            var headers = rows.First().Cells();

            var idxClass = headers.FirstOrDefault(x => !x.Value.IsBlank && x.Value.GetText() == HeaderClass)
                               ?.WorksheetColumn().ColumnNumber() ??
                           throw new InvalidOperationException($"Couldn't find a header named {HeaderClass}");
            var idxNode = headers.FirstOrDefault(x => !x.Value.IsBlank && x.Value.GetText() == HeaderNode)
                ?.WorksheetColumn().ColumnNumber() ??
                          throw new InvalidOperationException($"Couldn't find a header named {HeaderNode}");
            var idxRequiredMods = headers
                .FirstOrDefault(x => !x.Value.IsBlank && x.Value.GetText() == HeaderRequiredMods)
                ?.WorksheetColumn().ColumnNumber() ?? -1;
            var idxOriginal = headers.FirstOrDefault(x => !x.Value.IsBlank && x.Value.GetText() == HeaderOriginal)
                                  ?.WorksheetColumn().ColumnNumber() ??
                              throw new InvalidOperationException($"Couldn't find a header named {HeaderOriginal}");
            var idxTranslated = headers.FirstOrDefault(x => !x.Value.IsBlank && x.Value.GetText() == HeaderTranslated)
                ?.WorksheetColumn().ColumnNumber() ??
                                throw new InvalidOperationException($"Couldn't find a header named {HeaderTranslated}");



            for (int i = 1; i < rows.Count; i++)
            {
                var row = rows[i];
                var className = row.Cell(idxClass).Value.GetText();
                var node = row.Cell(idxNode).Value.GetText();
                List<string>? requiredMods = null;
                if (idxRequiredMods != -1 && row.Cell(idxRequiredMods).Value is { IsText: true } cellRequiredMods)
                {
                    requiredMods = cellRequiredMods.GetText().Split('\n').ToList();
                }
                var original = row.Cell(idxOriginal).Value.IsBlank ? "" : row.Cell(idxOriginal).Value.GetText();
                var cellTranslated = row.Cell(idxTranslated).Value;
                var translated = cellTranslated.IsText ? (cellTranslated.GetText() == "" ? null : cellTranslated.GetText()) : null;

                var translation = new TranslationEntry(className, node, original,
                    translated, requiredMods);
                translations.Add(translation);
            }
            return translations;
        }

        public static void ToLanguageXml(List<TranslationEntry> translations, bool skipNoTranslation, bool commentOriginal, string fileName, string rootDirPath)
        {
            var languagesDir = PathCombineCreateDir(rootDirPath, "Languages");
            var translationDir = PathCombineCreateDir(languagesDir, Prefabs.TranslationLanguage);
            var defInjected = new List<TranslationEntry>();
            var keyed = new List<TranslationEntry>();
            var strings = new List<TranslationEntry>();
            var patches = new List<TranslationEntry>();
            var patchedNodeSet = new HashSet<string>();

            Parallel.For(0, translations.Count, (i) =>
            {
                var originalTranslations = translations[i];
                var newTranslations = originalTranslations.DoNodeReplacement();
                if (ReferenceEquals(originalTranslations, newTranslations))
                    return;
                lock (translations)
                {
                    translations[i] = newTranslations;
                }
            });

            foreach (var translation in translations)
            {
                var className = translation.className;

                if (skipNoTranslation && className != "Strings" && string.IsNullOrEmpty(translation.translated))
                {
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
                        {
                            if (className.StartsWith("Patches."))
                                patches.Add(translation);
                            else
                                defInjected.Add(translation);
                            break;
                        }
                }
            }

            if (skipNoTranslation && patches.Count == 0 && defInjected.Count == 0 &&
                keyed.Count == 0 && translations.Count > 0)
            {
                Log.Wrn("번역 데이터가 존재하지 않아 아무것도 추출되지 않습니다.");
            }

            if (patches.Count > 0)
            {
                var outputPath = PathCombineCreateDir(rootDirPath, "Patches");

                var docPatch = new XmlDocument();
                docPatch.AppendChild(docPatch.CreateElement("Patch"));
                var root = docPatch.DocumentElement ?? throw new InvalidOperationException();

                // RequiredMods에 따라 뼈대 사전 생성
                foreach (var translation in patches)
                {
                    var requiredMods = translation.requiredMods;
                    if (requiredMods == null)
                        continue;
                    var a = root.ChildNodes.Where(x => x.HasAttribute("Class", "PatchOperationFindMod")).ToList();
                    foreach (var xmlNode in a)
                    {
                        var mods = xmlNode["mods"]?.ChildNodes.Select(y => y.InnerText).ToList();
                        var b =  mods?.HasSameElements(requiredMods) == true;
                    }
                    if (root.ChildNodes.Where(x => x.HasAttribute("Class", "PatchOperationFindMod"))
                        .Any(x =>
                        {
                            var mods = x["mods"]?.ChildNodes.Select(y => y.InnerText).ToList();
                            return mods?.HasSameElements(requiredMods) == true;
                        }))
                        continue;
                    root.AppendElement("Operation", operationFindMod =>
                    {
                        operationFindMod.AppendAttribute("Class", "PatchOperationFindMod");
                        operationFindMod.AppendElement("mods", mods =>
                        {
                            requiredMods.ForEach(requiredMod =>
                            {
                                if (requiredMod.StartsWith("##packageId##"))
                                {
                                    Log.ErrOnce(
                                        $"Required Mods 열에 잘못된 값이 존재합니다. Patches의 올바른 생성을 위해 엑셀 파일에 있는 해당 문구: \"{requiredMod}\" 를 직접 모드 이름으로 바꿔야 합니다.",
                                        $"잘못된{requiredMod}에러".GetHashCode());
                                }

                                mods.AppendElement("li", requiredMod);
                            });
                        });
                        operationFindMod.AppendElement("match", match =>
                        {
                            match.AppendAttribute("Class", "PatchOperationSequence");
                            match.AppendElement("success", "Always");
                            match.AppendElement("operations");
                        });
                        operationFindMod.AppendElement("nomatch", nomatch =>
                        {
                            nomatch.AppendAttribute("Class", "PatchOperationSequence");
                            nomatch.AppendElement("success", "Always");
                            nomatch.AppendElement("operations");
                        });
                    });


                }

                foreach (var translation in patches)
                {
                    var requiredMods = translation.requiredMods;
                    XmlElement operation;
                    if (requiredMods != null)
                    {
                        var operationFindMod = root.ChildNodes.FirstOrDefault(x =>
                            x.HasAttribute("Class", "PatchOperationFindMod") && 
                            x["mods"]!.ChildNodes.Select(x => x.InnerText)
                                .ToList().HasSameElements(requiredMods))!;

                        operation = operationFindMod["match"]!["operations"]!.AppendElement("li");


                        if (defInjected.Any(x => x.node == translation.node))
                        {
                            var noMatchTranslation = defInjected.First(x => x.node == translation.node);
                            operationFindMod["nomatch"]!["operations"]!.AppendElement("li", li =>
                            {
                                li.AppendAttribute("Class", "PatchOperationReplace");
                                li.AppendElement("success", "Always");
                                if (commentOriginal)
                                    li.AppendComment($"Original={SecurityElement.Escape(translation.original).Replace('-', 'ー')}");
                                li.AppendElement("xpath", Utils.GetXpath(translation.className[(translation.className.IndexOf('.') + 1)..], translation.node));
                                li.AppendElement("value", value =>
                                {
                                    var noMatchLastNode = translation.node.Split('.').Last();
                                    if (int.TryParse(noMatchLastNode, out _)) noMatchLastNode = "li";
                                    value.AppendElement(noMatchLastNode,
                                        noMatchTranslation.translated ?? noMatchTranslation.original);
                                });
                            });
                            patchedNodeSet.Add(translation.node);
                        }
                    }
                    else
                    {
                        operation = root.AppendElement("Operation");
                    }

                    operation.Append(li =>
                    {
                        li.AppendAttribute("Class", "PatchOperationReplace");
                        li.AppendElement("success", "Always");
                        if (commentOriginal)
                            li.AppendComment(
                                $"Original={SecurityElement.Escape(translation.original).Replace('-', 'ー')}");
                        li.AppendElement("xpath", Utils.GetXpath(translation.className[(translation.className.IndexOf('.') + 1)..], translation.node));
                        li.AppendElement("value", value =>
                        {
                            var lastNode = translation.node.Split('.').Last();
                            if (int.TryParse(lastNode, out _)) lastNode = "li";
                            value.AppendElement(lastNode, translation.translated ?? translation.original);
                        });
                    });

                }

                docPatch.SaveSafely(Path.Combine(outputPath, fileName + ".xml"));
            }

            if (defInjected.Count > 0)
            {
                var defInjectedDir = PathCombineCreateDir(translationDir, "DefInjected");
                var xmls = new Dictionary<string, XmlDocument>();
                foreach (var translation in defInjected)
                {
                    if (patchedNodeSet.Contains(translation.node))
                        continue;
                    PathCombineCreateDir(defInjectedDir, translation.className);
                    if (!xmls.TryGetValue(translation.className, out var doc))
                    {
                        doc = new XmlDocument();
                        xmls[translation.className] = doc;
                        doc.AppendChild(doc.CreateElement("LanguageData"));
                    }

                    doc.DocumentElement!.Append(languageData =>
                    {
                        if (commentOriginal)
                            languageData.AppendComment($"Original={SecurityElement.Escape(translation.original).Replace('-', 'ー')}");
                        languageData.AppendElement(translation.node, t =>
                        {
                            t.InnerText = translation.translated ?? translation.original;
                            if (!t.InnerText.Contains("{*")) return;
                            t.InnerText = Regex.Replace(t.InnerText, "\\{\\*(.*?)\\}", match =>
                            {
                                var targetIdentifier = match.Groups[1].Value;
                                var replacement = translations.FirstOrDefault(x => $"{x.className}+{x.node}" == targetIdentifier);
                                if (replacement != null)
                                    return replacement.translated ?? replacement.original;
                                Log.Err($"Pointer: {targetIdentifier}에 대한 원본 Identifier를 찾을 수 없습니다.");
                                return "ERR";
                            });
                        });
                    });

                }

                foreach (var (className, doc) in xmls)
                {
                    var outputPath = Path.Combine(defInjectedDir, className, fileName + ".xml");

                    doc.DoFullListTranslation();
                    doc.SaveSafely(outputPath);
                }
            }

            if (keyed.Count > 0)
            {
                var keyedDir = PathCombineCreateDir(translationDir, "Keyed");
                var xmls = new Dictionary<string, XmlDocument>();
                foreach (var translation in keyed)
                {
                    var idxSep = translation.node.IndexOf('|');
                    var key = idxSep != -1 ? translation.node.Split('|')[0] : "default";
                    var nodeName = idxSep != -1 ? translation.node[(idxSep + 1)..] : translation.node;

                    if (!xmls.TryGetValue(key, out var doc))
                    {
                        doc = new XmlDocument();
                        xmls[key] = doc;
                        doc.AppendChild(doc.CreateElement("LanguageData"));
                    }

                    doc.DocumentElement!.Append(languageData =>
                    {
                        if (commentOriginal)
                            languageData.AppendComment($"{Prefabs.OriginalLanguage}={SecurityElement.Escape(translation.original).Replace('-', 'ー')}");
                        languageData.AppendElement(nodeName, translation.translated ?? translation.original);
                    });
                }

                foreach (var (_, doc) in xmls)
                {
                    var outputPath = Path.Combine(keyedDir, $"{fileName}.xml");
                    doc.SaveSafely(outputPath);
                }
            }

            if (strings.Count > 0)
            {
                var stringDir = PathCombineCreateDir(translationDir, "Strings");
                var txts = new Dictionary<string, List<string>>();
                foreach (var translation in strings)
                {
                    var className = translation.node[..translation.node.LastIndexOf('.')];
                    if (!txts.TryGetValue(className, out var lines))
                    {
                        lines = new List<string>();
                        txts[className] = lines;
                    }

                    lines.Add(translation.translated ?? translation.original);
                }

                foreach (var (className, lines) in txts)
                {
                    var key = className[..className.LastIndexOf('.')].Replace('.', '\\');
                    var outputPath = PathCombineCreateDir(stringDir, key);
                    var fileNameTxt = Path.Combine(outputPath, $"{className.Split('.').Last()}") + ".txt";
                    lines.SaveSafely(fileNameTxt);
                }
            }
        }

        public static List<TranslationEntry> FromLanguageXml(string rootPath)
        {
            throw new NotImplementedException();
        }

        private static void SaveSafely(this XLWorkbook xlsx, string path)
        {
            if (!File.Exists(path))
            {
                xlsx.SaveAs(path);
                return;
            }

            switch (Prefabs.Policy)
            {
                case Prefabs.DuplicatesPolicy.Stop:
                    var stopCallback = Prefabs.StopCallbackXlsx;
                    if (stopCallback != null)
                        stopCallback(xlsx, path);
                    else
                        throw new ArgumentNullException(nameof(stopCallback));
                    return;
                case Prefabs.DuplicatesPolicy.Overwrite:
                    try
                    {
                        xlsx.SaveAs(path);
                    }
                    catch (IOException)
                    {
                        Log.Err("파일이 이미 사용 중이기 때문에 파일을 저장할 수 없었습니다.");
                    }
                    return;
                case Prefabs.DuplicatesPolicy.KeepOriginal:
                    return;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static void SaveSafely(this XmlDocument doc, string path)
        {
            doc.InsertBefore(doc.CreateXmlDeclaration("1.0", "utf-8", null), doc.DocumentElement);
            if (!File.Exists(path))
            {
                doc.Save(path);
                return;
            }

            switch (Prefabs.Policy)
            {
                case Prefabs.DuplicatesPolicy.Stop:
                    var stopCallback = Prefabs.StopCallbackXml;
                    if (stopCallback != null)
                        stopCallback(doc, path);
                    else
                        throw new ArgumentNullException(nameof(stopCallback));
                    return;
                case Prefabs.DuplicatesPolicy.Overwrite:
                    doc.Save(path);
                    return;
                case Prefabs.DuplicatesPolicy.KeepOriginal:
                    return;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static void SaveSafely(this IEnumerable<string> lines, string path)
        {
            if (!File.Exists(path))
            {
                File.WriteAllLines(path, lines);
                return;
            }

            switch (Prefabs.Policy)
            {
                case Prefabs.DuplicatesPolicy.Stop:
                    var stopCallback = Prefabs.StopCallbackTxt;
                    if (stopCallback != null)
                        stopCallback(lines, path);
                    else
                        throw new ArgumentNullException(nameof(stopCallback));
                    return;
                case Prefabs.DuplicatesPolicy.Overwrite:
                    File.WriteAllLines(path, lines);
                    return;
                case Prefabs.DuplicatesPolicy.KeepOriginal:
                    return;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
        private static string PathCombineCreateDir(params string[] paths)
        {
            var dir = Path.Combine(paths);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return dir;
        }

        public static string StripInvaildChars(this string str)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                str = str.Replace(c, ' ');
            }
            return str;
        }

        private static TranslationEntry DoNodeReplacement(this TranslationEntry orig)
        {
            if (orig.className is "Keyed" or "Strings")
                return orig;

            var isPatches = orig.className.StartsWith("Patches.");
            var defType = isPatches ? orig.className[("Patches.".Length + 1)..] : orig.className;
            var defName = orig.node.Split('.')[0];
            var nodeAfterDefName = orig.node[(orig.node.IndexOf('.') + 1)..];

            foreach (var (key, value) in Prefabs.NodeReplacement)
            {
                var tokenKey = key.Split('+');
                var tokenValue = value.Split("+");
                var targetDef = tokenKey[0];
                var targetNode = tokenKey[1];
                var changedDef = tokenValue[0];
                var changedNode = tokenValue[1];
                if (defType == targetDef && nodeAfterDefName == targetNode)
                {
                    return orig with { className = isPatches ? $"Patches.{changedDef}" : changedDef, node = $"{defName}.{changedNode}" };
                }
            }
            return orig;
        }

        private static void DoFullListTranslation(this XmlDocument defInjectedDoc)
        {
            var patterns = Prefabs.FullListTranslationTags.Select(x => $".+?\\.{x}\\.\\d+").ToList();

            var fullListdic = new Dictionary<string, XmlNode>();
            var removedNodesDic = new Dictionary<string, List<XmlNode>>();
            foreach (XmlNode childNode in defInjectedDoc.DocumentElement!.ChildNodes)
            {
                var nodeName = childNode.Name;
                if (!patterns.Any(x => Regex.IsMatch(nodeName, x)))
                    continue;
                nodeName = nodeName[..nodeName.LastIndexOf('.')];
                if (!fullListdic.TryGetValue(nodeName, out var fullList))
                {
                    fullList = defInjectedDoc.CreateElement(nodeName);
                    fullListdic[nodeName] = fullList;
                }

                if (!removedNodesDic.TryGetValue(nodeName, out var removedList))
                {
                    removedList = new List<XmlNode>();
                    removedNodesDic[nodeName] = removedList;
                }

                var li = fullList.AppendChild(defInjectedDoc.CreateElement("li"))!;
                li.InnerText = childNode.InnerText;
                removedList.Add(childNode);
            }


            foreach (var (key, fullListNode) in fullListdic)
            {
                var removedList = removedNodesDic[key];

                defInjectedDoc.DocumentElement!.InsertAfter(fullListNode, removedList.Last());
                foreach (var xmlNode in removedList)
                {
                    defInjectedDoc.DocumentElement!.RemoveChild(xmlNode);
                }
            }
        }
    }
}
