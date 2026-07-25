using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using RimworldExtractorInternal.Core;
using RimworldExtractorInternal.DataTypes;
using RimworldExtractorInternal.Procedures;

namespace RimworldExtractorInternal
{
    public static partial class Extractor
    {
        /// <summary>
        /// 완성된 SimulationResult로부터 번역 데이터(TranslationEntry)를 추출합니다.
        /// </summary>
        public static List<TranslationEntry> ExtractTranslationData(SimulationResult simResult)
        {
            bool isOfficialContent = simResult.TargetMod?.IsOfficialContent ?? false;
            var extraction = new List<TranslationEntry>();

            // 1. Defs 폴더 처리 (DefTree 시뮬레이션 결과 기반)
            var defsFolders = simResult.TargetFolders.Where(x => Path.GetFileName(x.FolderName) == "Defs").ToList();
            if (defsFolders.Count > 0)
            {
                extraction.AddRange(ExtractDefs(simResult, isOfficialContent));
            }

            // 2. 그 외 폴더 처리 (Keyed, Strings, Patches)
            foreach (var extractableFolder in simResult.TargetFolders)
            {
                switch (Path.GetFileName(extractableFolder.FolderName))
                {
                    case "Defs":
                        break;
                    case "Keyed":
                        extraction.AddRange(ExtractKeyed(extractableFolder, isOfficialContent));
                        break;
                    case "Strings":
                        extraction.AddRange(ExtractStrings(extractableFolder));
                        break;
                    case "Patches":
                        extraction.AddRange(ExtractPatches(simResult, extractableFolder, isOfficialContent));
                        break;
                    default:
                        Log.Wrn($"알 수 없는 추출 폴더입니다. {extractableFolder.FolderName}");
                        continue;
                }
            }

            // 3. 중복 노드 및 원문 불일치 검증
            var set = new HashSet<(string, string)>();
            foreach (var entry in extraction)
            {
                var tuple = (entry.ClassName + "+" + entry.Node, entry.Original);
                var pair = set.FirstOrDefault(x => x.Item1 == tuple.Item1);
                if (pair != default)
                {
                    if (pair.Item2 != entry.Original)
                    {
                        Log.Err($"중복 노드 원문 불일치 발견: {entry.ClassName}+{entry.Node}| {pair.Item2} | {entry.Original}");
                    }
                }
                set.Add(tuple);
            }

            return extraction.DistinctBy(x => $"{x.ClassName}+{x.Node}").ToList();
        }

        internal static IEnumerable<TranslationEntry> ExtractDefs(SimulationResult simResult, bool isOfficialContent)
        {
            var rawExtraction = ExtractDefsInternal(simResult, isOfficialContent).ToList();
#if DEBUG
            foreach (var translationEntry in rawExtraction)
            {
                Console.WriteLine(translationEntry);
            }
#endif
            // 🟢 ITranslationProcedure 기반의 외부 후처리 파이프라인 일괄 실행
            return TranslationPipelineRunner.Execute(rawExtraction);
        }

        private static IEnumerable<TranslationEntry> ExtractDefsInternal(SimulationResult simResult, bool isOfficialContent)
        {
            // 🟢 XML 전처리(DoPreProcessing)는 DefTreeSimulator 쪽으로 이관되었으므로 여기서는 순수 추출만 수행합니다.
            foreach (var node in simResult.DefTree.Root!.Elements()
                         .Where(x => x.Attribute("Reference")?.Value.ToLower() != "true"))
            {
                var defName = node.Element("defName")?.Value;
                if (defName == null)
                {
                    if (node.Name.LocalName != "SongDef")
                        Log.Wrn($"SongDef가 아닌 Abstract가 아닌 XML 요소에 'defName'이 없습니다. InnerXml: {node}");
                    continue;
                }

                var requiredMods = new RequiredMods();
                var requiredPackageIds = node.Attribute("RequiredPackageId")?.Value.Split(',');
                if (requiredPackageIds != null)
                {
                    requiredMods.AddAllowedByPackageIds(requiredPackageIds);
                }

                var className = node.Attribute("Class")?.Value ?? node.Name.LocalName;
                className = className[..1].ToUpper() + className[1..];

                foreach (var translationEntry in FindExtractableNodes(defName, className, node, isOfficialContent))
                {
                    yield return translationEntry with
                    {
                        RequiredMods = translationEntry.RequiredMods + requiredMods
                    };
                }
            }
        }

        internal static IEnumerable<TranslationEntry> ExtractKeyed(ExtractableFolder keyed, bool isOfficialContent)
        {
            var keyedRoot = keyed.FullPath;
            RequiredMods? requiredMods = keyed.RequiredPackageId == null
                ? null
                : new RequiredMods().Tap(rm => rm.AddAllowedByPackageIds(keyed.RequiredPackageId.Split(',')));

            foreach (var filePath in IO.DescendantFiles(keyedRoot).Where(x => x.ToLower().EndsWith(".xml")))
            {
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var doc = IO.ReadXml(filePath);

                foreach (var node in doc.Root!.Elements())
                {
                    yield return new TranslationEntry("Keyed", node.Name.LocalName, node.Value, null, requiredMods,
                        isOfficialContent ? fileName : null);
                }
            }
        }

        internal static IEnumerable<TranslationEntry> ExtractStrings(ExtractableFolder strings)
        {
            var stringsRoot = strings.FullPath;
            RequiredMods? requiredMods = strings.RequiredPackageId == null
                ? null
                : new RequiredMods().Tap(rm => rm.AddAllowedByPackageIds(strings.RequiredPackageId.Split(',')));

            foreach (var filePath in IO.DescendantFiles(stringsRoot).Where(x => x.ToLower().EndsWith(".txt")))
            {
                var nodeName = Path.GetRelativePath(stringsRoot, filePath);
                nodeName = Path.GetFileNameWithoutExtension(nodeName.Replace('\\', '.'));

                var lines = File.ReadAllLines(filePath);
                for (var i = 0; i < lines.Length; i++)
                {
                    var line = lines[i];
                    yield return new TranslationEntry("Strings", $"{nodeName}.{i}", line, null, requiredMods, null);
                }
            }
        }

        internal static IEnumerable<TranslationEntry> ExtractPatches(SimulationResult simResult, ExtractableFolder patches, bool isOfficialContent)
        {
            var rawExtraction = ExtractPatchesInternal(simResult, patches, isOfficialContent).ToList();
            
            // 🟢 ITranslationProcedure 기반의 외부 후처리 파이프라인 일괄 실행
            return TranslationPipelineRunner.Execute(rawExtraction);
        }

        private static IEnumerable<TranslationEntry> ExtractPatchesInternal(SimulationResult simResult, ExtractableFolder patches, bool isOfficialContent)
        {
            simResult.DefsAddedByPatches.Clear();
            var patchesRoot = patches.FullPath;
            RequiredMods? requiredMods = patches.RequiredPackageId == null
                ? null
                : new RequiredMods().Tap(rm => rm.AddAllowedByPackageIds(patches.RequiredPackageId.Split(',')));

            var doc = new XDocument(new XElement("Patch"));
            foreach (var filePath in IO.DescendantFiles(patchesRoot).Where(x => x.ToLower().EndsWith(".xml")))
            {
                var childDoc = IO.ReadXml(filePath);
                foreach (var node in childDoc.Root!.Elements())
                {
                    if (node.Name.LocalName != "Operation")
                        continue;
                    doc.Root!.Add(new XElement(node));
                }
            }

            foreach (var node in doc.Root!.Elements())
            {
                foreach (var translationEntry in PatchOperations.PatchOperationRecursive(node, simResult, null, false))
                {
                    yield return translationEntry with
                    {
                        RequiredMods = translationEntry.RequiredMods + requiredMods
                    };
                }
            }

            if (simResult.DefsAddedByPatches.Count == 0)
                yield break;

            foreach (var (requiredModsPatches, node) in simResult.DefsAddedByPatches)
            {
                var name = node.Attribute("Name")?.Value;
                if (requiredModsPatches != null)
                {
                    node.AppendElement("REQUIREDMODS", requiredModsPatches.ToString());
                }
                if (name != null)
                {
                    simResult.ParentNodeLookUp[name] = node;
                }
            }

            DefTreeSimulator.DoXmlInheritance(simResult, simResult.DefsAddedByPatches.Select(x => x.Item2));

            foreach (var translation in ExtractDefs(simResult, isOfficialContent))
            {
                yield return translation with
                {
                    ClassName = $"Patches.{translation.ClassName}",
                    RequiredMods = translation.RequiredMods + requiredMods
                };
            }
        }

        private static T Tap<T>(this T obj, Action<T> action) { action(obj); return obj; }
    }
}