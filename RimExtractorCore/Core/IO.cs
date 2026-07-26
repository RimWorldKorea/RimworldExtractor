using ClosedXML.Excel;
using System.Security;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RimExtractorCore
{
    public static class IO
    {
        private static readonly string HeaderClassNode = "Class+Node [(Identifier (Key)]";
        private static readonly string HeaderClass = "Class [Not chosen]";
        private static readonly string HeaderNode = "Node [Not chosen]";
        private static readonly string HeaderRequiredMods = "Required Mods [Not chosen]";
        private static string HeaderOriginal => $"{ConfigManager.Current.OriginalLanguage} [Source string]";
        private static string HeaderTranslated => $"{ConfigManager.Current.TranslationLanguage} [Translation]";

        public static void ToExcel(List<TranslationEntry> translations, string outPath = "result",
            bool markNoTranslation = false)
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
                sheet.Cell(2 + i, 1).Value = $"{entry.ClassName}+{entry.Node}";
                sheet.Cell(2 + i, 2).Value = entry.ClassName;
                sheet.Cell(2 + i, 3).Value = entry.Node;
                if (entry.RequiredMods != null)
                {
                    var combinedRequiredMods = entry.RequiredMods.ToString();
                    var cellRequiredMods = sheet.Cell(2 + i, 4);
                    cellRequiredMods.Value = combinedRequiredMods;
                    if (combinedRequiredMods.Contains("##packageId##") && entry.ClassName.StartsWith("Patches."))
                    {
                        Log.WrnOnce($"Required Mods 열에 잘못된 값이 존재합니다. 추후 Patches의 올바른 생성을 위해 엑셀 파일에 있는 해당 문구: \"{combinedRequiredMods}\" 를 직접 모드 이름으로 바꿔야 합니다.",
                            $"잘못된{combinedRequiredMods}경고".GetHashCode());
                        var comment = cellRequiredMods.GetComment();
                        comment.AddText(
                            $"모드 이름 대신 패키지 이름({RequiredMods.PACKAGE_ID_PREFIX})이 있습니다. 모드 이름으로 올바르게 수정해주세요.");
                        comment.Visible = true;
                    }
                }
                sheet.Cell(2 + i, 5).Value = entry.Original;
                if (entry.Translated != null)
                {
                    sheet.Cell(2 + i, 6).Value = entry.Translated;
                }
                else if (markNoTranslation)
                {
                    sheet.Cell(2 + i, 6).Style.Fill.SetBackgroundColor(XLColor.SkyBlue);
                }

                if (entry.TryGetExtension(ExtractorConstants.ExtensionKeyExtraCommentTranslated, out object? extension) &&
                    extension is string extensionStr)
                {
                    var comment = sheet.Cell(2 + i, 6).CreateComment();
                    comment.AddText(extensionStr);
                    comment.Visible = true;
                }
            }

            sheet.Style.Font.FontName = "맑은 고딕";
            xlsx.SaveSafely(outPath + ".xlsx");
        }

        public static void ModifyExcel(List<TranslationAnalyzerEntry.ChangeRecord> changes, string targetPath)
        {
            var xlsx = new XLWorkbook(targetPath);
            var mainSheet = xlsx.Worksheets.Worksheet(1);
            var rows = mainSheet.RowsUsed().ToList();
            var offset = rows.Count;
            var headers = rows.First().Cells();

            var colClassNode =
                headers.FirstOrDefault(x => x.StrVal() == HeaderClassNode)?.WorksheetColumn().ColumnNumber() ??
                throw new XlsxHeaderReadingException(HeaderClass);
            var colClass = headers.FirstOrDefault(x => x.StrVal() == HeaderClass)
                               ?.WorksheetColumn().ColumnNumber() ??
                           throw new XlsxHeaderReadingException(HeaderClass);
            var colNode = headers.FirstOrDefault(x => x.StrVal() == HeaderNode)
                              ?.WorksheetColumn().ColumnNumber() ??
                          throw new XlsxHeaderReadingException(HeaderNode);
            var colRequiredMods = headers
                .FirstOrDefault(x => x.StrVal() == HeaderRequiredMods)
                ?.WorksheetColumn().ColumnNumber() ?? -1;
            var colOriginal = headers.FirstOrDefault(x => x.StrVal() == HeaderOriginal)
                                  ?.WorksheetColumn().ColumnNumber() ??
                              headers.FirstOrDefault(x => x.StrVal() == "EN [Source string]")
                                  ?.WorksheetColumn().ColumnNumber() ??
                              throw new XlsxHeaderReadingException(HeaderOriginal);
            var colTranslated = headers.FirstOrDefault(x => x.StrVal() == HeaderTranslated)
                                    ?.WorksheetColumn().ColumnNumber() ??
                                headers.FirstOrDefault(x => x.StrVal() == "KO [Translation]")
                                    ?.WorksheetColumn().ColumnNumber() ??
                                throw new XlsxHeaderReadingException(HeaderTranslated);
            
            var changedOriginals = new List<TranslationAnalyzerEntry.ChangeRecord>();
            var fillOriginals = new List<TranslationAnalyzerEntry.ChangeRecord>();
            var removeNodes = new List<TranslationAnalyzerEntry.ChangeRecord>();
            var addedNewlys = new List<TranslationAnalyzerEntry.ChangeRecord>();
            var dateString = DateTime.Today.ToString("yyyy-MM-dd");

            foreach (var changeRecord in changes)
            {
                switch (changeRecord.Reason)
                {
                    case TranslationAnalyzerEntry.ChangeReason.ChangedOriginal:
                        changedOriginals.Add(changeRecord);
                        break;
                    case TranslationAnalyzerEntry.ChangeReason.FillOriginal:
                        fillOriginals.Add(changeRecord);
                        break;
                    case TranslationAnalyzerEntry.ChangeReason.RemoveNode:
                        removeNodes.Add(changeRecord);
                        break;
                    case TranslationAnalyzerEntry.ChangeReason.AddedNewly:
                        addedNewlys.Add(changeRecord);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }

            if (removeNodes.Count > 0)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    var curRow = rows[i];
                    var curClassNode = curRow.Cell(colClassNode).StrVal();
                    var pairEntry =
                        removeNodes.FirstOrDefault(x => $"{x.Orig!.ClassName}+{x.Orig!.Node}" == curClassNode);
                    if (pairEntry != null)
                    {
                        var origCell = curRow.Cell(colOriginal);
                        origCell.GetComment().AddText($"{dateString}에 삭제됨. 삭제 이전 번역문: '{curRow.Cell(colTranslated).StrVal()}'\n");
                        origCell.GetComment().Visible = true;
                        origCell.Style.Fill.SetBackgroundColor(XLColor.Red);
                        curRow.Cell(colTranslated).Clear();
                    }
                }
            }

            if (changedOriginals.Count > 0)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    var curRow = rows[i];
                    var curClassNode = curRow.Cell(colClassNode).StrVal();
                    var pairEntry =
                        changedOriginals.FirstOrDefault(x => $"{x.Orig!.ClassName}+{x.Orig!.Node}" == curClassNode);
                    if (pairEntry != null)
                    {
                        var origCell = curRow.Cell(colOriginal);
                        origCell.GetComment().AddText($"{dateString} 이전의 원문: '{curRow.Cell(colOriginal).StrVal()}'\n");
                        origCell.Value = pairEntry.New!.Original;
                        origCell.GetComment().Visible = true;
                        origCell.Style.Fill.SetBackgroundColor(XLColor.Orange);
                    }
                }
            }

            if (fillOriginals.Count > 0)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    var curRow = rows[i];
                    var curClassNode = curRow.Cell(colClassNode).StrVal();
                    var pairEntry =
                        fillOriginals.FirstOrDefault(x => $"{x.Orig!.ClassName}+{x.Orig!.Node}" == curClassNode);
                    if (pairEntry != null)
                    {
                        var origCell = curRow.Cell(colOriginal);
                        origCell.GetComment()
                            .AddText($"{dateString}에 소실되었던 원문이 추가되었습니다.\n");
                        origCell.Value = pairEntry.New!.Original;
                        origCell.GetComment().Visible = true;
                        origCell.Style.Fill.SetBackgroundColor(XLColor.Orange);
                    }
                }
            }

            if (addedNewlys.Count > 0)
            {
                for (int i = 0; i < addedNewlys.Count; i++)
                {
                    var entry = addedNewlys[i].New;
                    mainSheet.Cell(2 + i + rows.Count, colClassNode).Value = $"{entry.ClassName}+{entry.Node}";
                    mainSheet.Cell(2 + i + rows.Count, colClass).Value = entry.ClassName;
                    mainSheet.Cell(2 + i + rows.Count, colNode).Value = entry.Node;
                    if (colRequiredMods != -1 && entry.RequiredMods != null)
                    {
                        var combinedRequiredMods = entry.RequiredMods.ToString();
                        mainSheet.Cell(2 + i + rows.Count, colRequiredMods).Value = combinedRequiredMods;
                        if (combinedRequiredMods.Contains("##packageId##") && entry.ClassName.StartsWith("Patches."))
                        {
                            Log.WrnOnce($"Required Mods 열에 잘못된 값이 존재합니다. 추후 Patches의 올바른 생성을 위해 엑셀 파일에 있는 해당 문구: \"{combinedRequiredMods}\" 를 직접 모드 이름으로 바꿔야 합니다.",
                                $"잘못된{combinedRequiredMods}경고".GetHashCode());
                        }
                    }
                    mainSheet.Cell(2 + i + rows.Count, colOriginal).Value = entry.Original;
                    if (entry.Translated != null)
                    {
                        mainSheet.Cell(2 + i + rows.Count, colTranslated).Value = entry.Translated;
                    }

                    if (entry.TryGetExtension(ExtractorConstants.ExtensionKeyExtraCommentTranslated, out object? extension) &&
                        extension is string extensionStr)
                    {
                        var comment = mainSheet.Cell(2 + i + rows.Count, 6).GetComment();
                        comment.AddText(extensionStr);
                        comment.Visible = true;
                    }
                    mainSheet.Row(2 + i + rows.Count).Select();

                    if (i == 0)
                    {
                        mainSheet.Cell(2 + i + rows.Count, colOriginal).Style.Fill.SetBackgroundColor(XLColor.SkyBlue);
                        var comment = mainSheet.Cell(2 + i + rows.Count, colOriginal).GetComment();
                        comment.AddText($"{dateString}에 새로 추가된 노드들 ({addedNewlys.Count}개)");
                        comment.Visible = true;
                    }
                }
            }

            mainSheet.Style.Font.FontName = "맑은 고딕";
            foreach (var cell in mainSheet.CellsUsed().Where(x => x.HasComment))
            {
                var comment = cell.GetComment();
                comment.Position.ColumnOffset = 5d;
                comment.Position.RowOffset = 5d;
                comment.Position.Row = cell.Address.RowNumber + 1;
                comment.Position.Column = cell.Address.ColumnNumber + 1;
                comment.Style.Alignment.SetAutomaticSize();
            }
            xlsx.SaveSafely(targetPath);
        }

        public static List<TranslationEntry> FromExcel(string inputPath)
        {
            return Spreadsheet.SpreadsheetReader.ReadTranslations(inputPath);
        }

        public static void ToLanguageXml(List<TranslationEntry> translations, bool skipNoTranslation, bool commentOriginal, string ModName, string rootDirPath)
        {
            var languagesDir = PathCombineCreateDir(rootDirPath, "Languages");
            var translationDir = PathCombineCreateDir(languagesDir, ConfigManager.Current.TranslationLanguage);
            var defInjected = new List<TranslationEntry>();
            var defInjectedFullListTranslations = new List<TranslationEntry>();
            var keyed = new List<TranslationEntry>();
            var strings = new List<TranslationEntry>();
            
            var conditionalDefInjected = new List<TranslationEntry>();

            var isOfficial = translations.Any(x => x.SourceFile != null);
            if (isOfficial)
            {
                Log.Msg("공식 컨텐츠는 모드를 추출할 때와는 달리, 파일명을 보존해서 추출합니다.");
            }

            foreach (var rawTranslation in translations)
            {
                var translation = rawTranslation;
                var className = translation.ClassName;

                if (skipNoTranslation && className != "Strings" && string.IsNullOrEmpty(translation.Translated))
                {
                    continue;
                }

                if (className.StartsWith("Patches."))
                {
                    var realClassName = className["Patches.".Length..];
                    translation = translation with { ClassName = realClassName };

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
                        {
                            if (!isOfficial && ConfigManager.Current.FullListTranslationTags.Any(translation.Node.Contains))
                                defInjectedFullListTranslations.Add(translation);
                            else
                                defInjected.Add(translation);
                            break;
                        }
                }
            }

            if (skipNoTranslation && conditionalDefInjected.Count == 0 && defInjected.Count == 0 &&
                keyed.Count == 0 && translations.Count > 0 && defInjectedFullListTranslations.Count == 0)
            {
                Log.Wrn("번역 데이터가 존재하지 않아 아무것도 추출되지 않습니다. 팁) XLSX -> XML 기능의 경우 번역된 내용이 없으면 아무것도 저장되지 않습니다.");
            }

            if (conditionalDefInjected.Count > 0)
            {
                var conditionalBaseDir = PathCombineCreateDir(rootDirPath, "ConditionalLanguages");
                var groupedByMods = conditionalDefInjected.GroupBy(x => x.RequiredMods?.ToString() ?? "Common");

                foreach (var group in groupedByMods)
                {
                    var packageIdFolder = group.Key.Replace("::", "_").Replace("/", "_").Replace('\\', '_');
                    var targetFolder = PathCombineCreateDir(conditionalBaseDir, packageIdFolder, ConfigManager.Current.TranslationLanguage, "DefInjected");

                    var xmls = new Dictionary<string, XDocument>();
                    foreach (var translation in group)
                    {
                        PathCombineCreateDir(targetFolder, translation.ClassName);
                        var key = $"{translation.ClassName}|{translation.SourceFile}";

                        if (!xmls.TryGetValue(key, out var doc))
                        {
                            doc = new XDocument(new XElement("LanguageData"));
                            xmls[key] = doc;
                        }

                        if (commentOriginal)
                            doc.Root!.AppendComment($"Original={SecurityElement.Escape(translation.Original).Replace('-', 'ー')}");

                        var elem = doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
                        ProcessPointerReplacement(elem, translations);
                    }

                    foreach (var (key, doc) in xmls)
                    {
                        var tokens = key.Split('|');
                        var className = tokens[0];
                        var outputPath = isOfficial
                            ? Path.Combine(targetFolder, className, tokens[1] + ".xml")
                            : Path.Combine(targetFolder, className, Utils.GenerateFileName(Path.GetFileNameWithoutExtension(ModName), className) + ".xml");

                        doc.DoFullListTranslation();
                        doc.SaveSafely(outputPath);
                    }
                }
            }

            if (defInjected.Count > 0)
            {
                var defInjectedDir = PathCombineCreateDir(translationDir, "DefInjected");
                var xmls = new Dictionary<string, XDocument>();

                foreach (var translation in defInjected)
                {
                    PathCombineCreateDir(defInjectedDir, translation.ClassName);
                    var key = $"{translation.ClassName}|{translation.SourceFile}";

                    if (!xmls.TryGetValue(key, out var doc))
                    {
                        doc = new XDocument(new XElement("LanguageData"));
                        xmls[key] = doc;
                    }

                    if (commentOriginal)
                        doc.Root!.AppendComment($"Original={SecurityElement.Escape(translation.Original).Replace('-', 'ー')}");

                    var elem = doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
                    ProcessPointerReplacement(elem, translations);
                }

                foreach (var (key, doc) in xmls)
                {
                    var tokens = key.Split('|');
                    var className = tokens[0];
                    var outputPath = isOfficial
                        ? Path.Combine(defInjectedDir, className, tokens[1] + ".xml")
                        : Path.Combine(defInjectedDir, className, Utils.GenerateFileName(Path.GetFileNameWithoutExtension(ModName), className) + ".xml");

                    doc.DoFullListTranslation();
                    doc.SaveSafely(outputPath);
                }
            }

            if (defInjectedFullListTranslations.Count > 0)
            {
                var defInjectedDir = PathCombineCreateDir(translationDir, "DefInjected");
                var xmls = new Dictionary<(string, string), XDocument>();

                foreach (var translation in defInjectedFullListTranslations)
                {
                    PathCombineCreateDir(defInjectedDir, translation.ClassName);
                    var nodeParent = translation.Node[..translation.Node.LastIndexOf('.')];
                    var key = (translation.ClassName, nodeParent);

                    if (!xmls.TryGetValue(key, out var doc))
                    {
                        doc = new XDocument(new XElement("LanguageData"));
                        xmls[key] = doc;
                    }

                    if (commentOriginal)
                        doc.Root!.AppendComment($"Original={SecurityElement.Escape(translation.Original).Replace('-', 'ー')}");

                    var elem = doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
                    ProcessPointerReplacement(elem, translations);
                }

                foreach (var ((className, nodeParent), doc) in xmls)
                {
                    var outputPath = Path.Combine(defInjectedDir, className,
                        Utils.GenerateFileName(Path.GetFileNameWithoutExtension(ModName), className, nodeParent) + ".xml");

                    doc.DoFullListTranslation();
                    doc.SaveSafely(outputPath);
                }
            }

            if (keyed.Count > 0)
            {
                var keyedDir = PathCombineCreateDir(translationDir, "Keyed");
                var xmls = new Dictionary<string, XDocument>();

                foreach (var translation in keyed)
                {
                    var key = isOfficial ? translation.SourceFile! : "default";

                    if (!xmls.TryGetValue(key, out var doc))
                    {
                        doc = new XDocument(new XElement("LanguageData"));
                        xmls[key] = doc;
                    }

                    if (commentOriginal)
                        doc.Root!.AppendComment($"{ConfigManager.Current.OriginalLanguage}={SecurityElement.Escape(translation.Original).Replace('-', 'ー')}");

                    doc.Root!.AppendElement(translation.Node, translation.Translated ?? translation.Original);
                }

                foreach (var (key, doc) in xmls)
                {
                    var outputPath = isOfficial 
                        ? Path.Combine(keyedDir, $"{key}.xml")
                        : Path.Combine(keyedDir, Utils.GenerateFileName(Path.GetFileNameWithoutExtension(ModName), "Keyed") + ".xml");
                    
                    doc.SaveSafely(outputPath);
                }
            }

            if (strings.Count > 0)
            {
                var stringDir = PathCombineCreateDir(translationDir, "Strings");
                var txts = new Dictionary<string, List<string>>();

                foreach (var translation in strings)
                {
                    var className = translation.Node[..translation.Node.LastIndexOf('.')];
                    if (!txts.TryGetValue(className, out var lines))
                    {
                        lines = new List<string>();
                        txts[className] = lines;
                    }

                    lines.Add(translation.Translated ?? translation.Original);
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

        private static void ProcessPointerReplacement(XElement elem, List<TranslationEntry> translations)
        {
            if (!elem.Value.Contains("{*")) return;
            elem.Value = Regex.Replace(elem.Value, "\\{\\*(.*?)\\}", match =>
            {
                var targetIdentifier = match.Groups[1].Value;
                var replacement = translations.FirstOrDefault(x => $"{x.ClassName}+{x.Node}" == targetIdentifier);
                if (replacement != null)
                    return replacement.Translated ?? replacement.Original;
                Log.Err($"Pointer: {targetIdentifier}에 대한 원본 Identifier를 찾을 수 없습니다.");
                return "ERR";
            });
        }

        public static List<TranslationEntry> FromLanguageXml(string rootPath, bool isOfficialContent = false)
        {
            var translationsDir = Path.Combine(rootPath, "Languages", ConfigManager.Current.TranslationLanguage);
            if (!Directory.Exists(translationsDir))
                translationsDir = Path.Combine(rootPath, "Languages", ConfigManager.Current.TranslationLanguage.Split(' ').First());

            var defInjectedDir = Path.Combine(translationsDir, "DefInjected");
            var keyedDir = Path.Combine(translationsDir, "Keyed");
            var stringsDir = Path.Combine(translationsDir, "Strings");

            var translations = new List<TranslationEntry>();

            foreach (var filePath in DescendantFiles(defInjectedDir).Where(x => x.ToLower().EndsWith(".xml")))
            {
                var className = Path.GetRelativePath(defInjectedDir, filePath).Split(Path.DirectorySeparatorChar)
                    .First();
                try
                {
                    var doc = ReadXml(filePath);
                    foreach (var node in doc.Root!.Elements())
                    {
                        var name = node.Name.LocalName;
                        if (node.Elements().Any())
                        {
                            var children = node.Elements().ToList();
                            for (int i = 0; i < children.Count; i++)
                            {
                                translations.Add(new TranslationEntry(className, $"{name}.{i}", string.Empty,
                                    children[i].Value, null, null));
                            }
                        }
                        else
                        {
                            translations.Add(
                                new TranslationEntry(className, name, string.Empty, node.Value, null, null));
                        }
                    }
                }
                catch (Exception e)
                {
                    Log.Err($"{filePath}를 읽는 중 에러 발생: {e.Message}");
                    throw;
                }
            }

            // 🟢 bool isOfficialContent 인수 전달
            var keyed = new ExtractableFolder(ModMetadata.Emptry, keyedDir, null);
            translations.AddRange(Extractor.ExtractKeyed(keyed, isOfficialContent)
                .Select(x => x with { Translated = x.Original, Original = "" }));

            var strings = new ExtractableFolder(ModMetadata.Emptry, stringsDir, null);
            translations.AddRange(Extractor.ExtractStrings(strings)
                .Select(x => x with { Translated = x.Original, Original = "" }));

            return translations;
        }

        private static void SaveSafely(this XLWorkbook xlsx, string path)
        {
            if (!File.Exists(path))
            {
                xlsx.SaveAs(path);
                return;
            }

            switch (ConfigManager.Current.Policy)
            {
                case DuplicatesPolicy.Stop:
                    var stopCallback = ExtractorConstants.StopCallbackXlsx;
                    if (stopCallback != null)
                        stopCallback(xlsx, path);
                    else
                        throw new ArgumentNullException(nameof(stopCallback));
                    return;
                case DuplicatesPolicy.Overwrite:
                    try
                    {
                        xlsx.SaveAs(path);
                    }
                    catch (IOException)
                    {
                        Log.Err($"{Path.GetFileName(path)}: 파일이 이미 사용 중이기 때문에 파일을 저장할 수 없었습니다. 종료 후 재시도 해주세요.");
                    }
                    return;
                case DuplicatesPolicy.KeepOriginal:
                    return;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static void SaveSafely(this XDocument doc, string path)
        {
            if (!File.Exists(path))
            {
                doc.Save(path);
                return;
            }

            switch (ConfigManager.Current.Policy)
            {
                case DuplicatesPolicy.Stop:
                    var stopCallback = ExtractorConstants.StopCallbackXml;
                    if (stopCallback != null)
                        stopCallback(doc, path);
                    else
                        throw new ArgumentNullException(nameof(stopCallback));
                    return;
                case DuplicatesPolicy.Overwrite:
                    doc.Save(path);
                    return;
                case DuplicatesPolicy.KeepOriginal:
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

            switch (ConfigManager.Current.Policy)
            {
                case DuplicatesPolicy.Stop:
                    var stopCallback = ExtractorConstants.StopCallbackTxt;
                    if (stopCallback != null)
                        stopCallback(lines, path);
                    else
                        throw new ArgumentNullException(nameof(stopCallback));
                    return;
                case DuplicatesPolicy.Overwrite:
                    File.WriteAllLines(path, lines);
                    return;
                case DuplicatesPolicy.KeepOriginal:
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

        private static void DoFullListTranslation(this XDocument defInjectedDoc)
        {
            var patterns = ConfigManager.Current.FullListTranslationTags.Select(x => $".+?\\.{x}\\.\\d+").ToList();

            var fullListdic = new Dictionary<string, XElement>();
            var removedNodesDic = new Dictionary<string, List<XElement>>();
            foreach (var childNode in defInjectedDoc.Root!.Elements().ToList())
            {
                var nodeName = childNode.Name.LocalName;
                if (!patterns.Any(x => Regex.IsMatch(nodeName, x)))
                    continue;
                nodeName = nodeName[..nodeName.LastIndexOf('.')];
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

                var li = fullList.AppendElement("li", childNode.Value);
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

        internal static XDocument ReadXml(string filePath)
        {
            var readerSettings = new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true,
                CheckCharacters = false
            };
    
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var xmlReader = XmlReader.Create(stream, readerSettings);
            return XDocument.Load(xmlReader);
        }

        internal static IEnumerable<string> DescendantFiles(string root)
        {
            if (!Directory.Exists(root))
                yield break;

            var q = new Queue<string>();
            q.Enqueue(root);

            while (q.Count > 0)
            {
                var curPath = q.Dequeue();
                foreach (var subDir in Directory.GetDirectories(curPath).OrderBy(x => x))
                {
                    q.Enqueue(subDir);
                }

                foreach (var file in Directory.EnumerateFiles(curPath).OrderBy(x => x))
                {
                    yield return file;
                }
            }
        }
    }
}