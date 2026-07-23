using System.Xml.Linq;
using RimworldExtractorInternal.Compats;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorInternal
{
    public static partial class Extractor
    {
        private static bool _isOfficialContent = false;

        public static List<TranslationEntry> ExtractTranslationData(ModMetadata modMetadata, List<ExtractableFolder> selectedFolders, List<ModMetadata>? referenceMods)
        {
            if (modMetadata.IsOfficialContent)
                _isOfficialContent = true;

            var refDefs = new List<string>();
            var prePatches = new List<ExtractableFolder>();
            if (referenceMods != null)
            {
                foreach (var referenceMod in referenceMods)
                {
                    refDefs.AddRange(from extractableFolder in ModLister.GetExtractableFolders(referenceMod)
                        where (extractableFolder.VersionInfo == "default" ||
                               extractableFolder.VersionInfo == "Common" ||
                               extractableFolder.VersionInfo == Prefabs.CurrentVersion)
                              && Path.GetFileName(extractableFolder.FolderName) == "Defs"
                        select Path.Combine(referenceMod.RootDir, extractableFolder.FolderName));
                    prePatches.AddRange(ModLister.GetExtractableFolders(referenceMod).Where(x =>
                        (x.VersionInfo == "default" || x.VersionInfo == "Common" ||
                         x.VersionInfo == Prefabs.CurrentVersion) && Path.GetFileName(x.FolderName) == "Patches"));
                }
            }

            // 1. 환경 모사 실행 -> 최종 모사된 XDocument 수령
            var simResult = DefTreeSimulator.Execute(
                modMetadata, selectedFolders, prePatches, refDefs, _isOfficialContent);

            var extraction = new List<TranslationEntry>();

            // 2. 모사된 XDocument 트리를 순회하며 TranslationEntry 추출
            var defs = selectedFolders.Where(x => Path.GetFileName(x.FolderName) == "Defs").ToList();
            if (defs.Count > 0)
            {
                extraction.AddRange(ExtractDefs(simResult));
            }

            foreach (var extractableFolder in selectedFolders)
            {
                switch (Path.GetFileName(extractableFolder.FolderName))
                {
                    case "Defs":
                        break;
                    case "Keyed":
                        extraction.AddRange(ExtractKeyed(extractableFolder));
                        break;
                    case "Strings":
                        extraction.AddRange(ExtractStrings(extractableFolder));
                        break;
                    case "Patches":
                        extraction.AddRange(ExtractPatches(simResult, extractableFolder));
                        break;
                    default:
                        Log.Wrn($"지원하지 않는 폴더입니다. {extractableFolder.FolderName}");
                        continue;
                }
            }

            var set = new HashSet<(string, string)>();
            foreach (var entry in extraction)
            {
                var tuple = (entry.ClassName + "+" + entry.Node, entry.Original);
                var pair = set.FirstOrDefault(x => x.Item1 == tuple.Item1);
                if (pair != default)
                {
                    if (pair.Item2 != entry.Original)
                    {
                        Log.Err(
                            $"원문이 다른 중복되는 노드가 있습니다. 노드: {entry.ClassName}+{entry.Node}| {pair.Item2} | {entry.Original} ");
                    }
                }

                set.Add(tuple);
            }

            _isOfficialContent = false;

            return extraction.DistinctBy(x => $"{x.ClassName}+{x.Node}").ToList();
        }

        internal static IEnumerable<TranslationEntry> ExtractDefs(SimulationResult simResult)
        {
            var rawExtraction = ExtractDefsInternal(simResult).ToList();
            foreach (var translationEntry in rawExtraction)
            {
                Console.WriteLine(translationEntry);
            }
            foreach (var entry in CompatManager.DoPostProcessing(rawExtraction))
            {
                yield return entry;
            }
        }

        private static IEnumerable<TranslationEntry> ExtractDefsInternal(SimulationResult simResult)
        {
            CompatManager.DoPreProcessing(simResult.DefTree);

            foreach (var node in simResult.DefTree.Root!.Elements()
                         .Where(x => x.Attribute("Reference")?.Value.ToLower() != "true"))
            {
                var defName = node.Element("defName")?.Value;
                if (defName == null)
                {
                    if (node.Name.LocalName != "SongDef")
                        Log.Wrn($"SongDef과 Abstract가 아닌 XML 요소 {node.Name}에서 'defName' 태그를 찾지 못했습니다. InnerXml: {node}");
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

                foreach (var translationEntry in FindExtractableNodes(defName, className, node))
                {
                    yield return translationEntry with
                    {
                        RequiredMods = translationEntry.RequiredMods + requiredMods
                    };
                }
            }
        }

        internal static IEnumerable<TranslationEntry> ExtractKeyed(ExtractableFolder keyed)
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
                        _isOfficialContent ? fileName : null);
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

        internal static IEnumerable<TranslationEntry> ExtractPatches(SimulationResult simResult, ExtractableFolder patches)
        {
            var rawExtraction = ExtractPatchesInternal(simResult, patches).ToList();
            foreach (var entry in CompatManager.DoPostProcessing(rawExtraction))
            {
                yield return entry;
            }
        }

        private static IEnumerable<TranslationEntry> ExtractPatchesInternal(SimulationResult simResult, ExtractableFolder patches)
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

            CompatManager.DoPreProcessing(doc);
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

            foreach (var translation in ExtractDefs(simResult))
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