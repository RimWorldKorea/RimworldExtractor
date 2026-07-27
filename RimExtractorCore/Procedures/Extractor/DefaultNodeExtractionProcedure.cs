using System.Xml.Linq;
using System.Collections;
using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;
using RimExtractorCore.Extractor;

namespace RimExtractorCore.Procedures;

public class DefaultNodeExtractionProcedure : IExtractionProcedure
{
    public string Name => "DefaultNodeExtractionProcedure";

    public IEnumerable<TranslationEntry> Extract(SimulationResult simResult)
    {
        bool isOfficialContent = simResult.TargetMod?.IsOfficialContent ?? false;
        var extraction = new List<TranslationEntry>();

        // 1. Defs 추출
        var defsFolders = simResult.TargetFolders.Where(x => Path.GetFileName(x.FolderName) == "Defs").ToList();
        if (defsFolders.Count > 0)
        {
            extraction.AddRange(ExtractDefs(simResult, isOfficialContent));
        }

        // 2. Keyed, Strings, Patches 추출
        foreach (var extractableFolder in simResult.TargetFolders)
        {
            switch (Path.GetFileName(extractableFolder.FolderName))
            {
                case "Defs": break;
                case "Keyed": extraction.AddRange(ExtractKeyed(extractableFolder, isOfficialContent)); break;
                case "Strings": extraction.AddRange(ExtractStrings(extractableFolder)); break;
                case "Patches": extraction.AddRange(ExtractPatches(simResult, extractableFolder, isOfficialContent)); break;
                default:
                    Log.Wrn($"알 수 없는 폴더입니다. {extractableFolder.FolderName}");
                    continue;
            }
        }

        return extraction;
    }

    internal IEnumerable<TranslationEntry> ExtractDefs(SimulationResult simResult, bool isOfficialContent)
    {
        foreach (var node in simResult.DefTree.Root!.Elements()
                     .Where(x => x.Attribute("Reference")?.Value.ToLower() != "true"))
        {
            var defName = node.Element("defName")?.Value;
            if (defName == null)
            {
                if (node.Name.LocalName != "SongDef")
                    Log.Wrn($"SongDef가 아닌 Abstract하지 않은 XML 노드에 'defName'이 없습니다. InnerXml: {node}");
                continue;
            }

            var requiredMods = new RequiredMods();
            var requiredPackageIds = node.Attribute("RequiredPackageId")?.Value.Split(',');
            if (requiredPackageIds != null)
                requiredMods.AddAllowedByPackageIds(requiredPackageIds);

            var className = node.Attribute("Class")?.Value ?? node.Name.LocalName;
            className = className[..1].ToUpper() + className[1..];

            foreach (var translationEntry in ExtractorEngine.FindExtractableNodes(defName, className, node, isOfficialContent))
            {
                yield return translationEntry with { RequiredMods = translationEntry.RequiredMods + requiredMods };
            }
        }
    }

    private IEnumerable<TranslationEntry> ExtractKeyed(ExtractableFolder keyed, bool isOfficialContent)
    {
        var keyedRoot = keyed.FullPath;
        RequiredMods? requiredMods = keyed.RequiredPackageId == null
            ? null : new RequiredMods().Tap(rm => rm.AddAllowedByPackageIds(keyed.RequiredPackageId.Split(',')));

        foreach (var filePath in FileInterface.DescendantFiles(keyedRoot).Where(x => x.ToLower().EndsWith(".xml")))
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var doc = FileInterface.ReadXml(filePath);
            foreach (var node in doc.Root!.Elements())
            {
                yield return new TranslationEntry("Keyed", node.Name.LocalName, node.Value, null, requiredMods,
                    isOfficialContent ? fileName : null);
            }
        }
    }

    private IEnumerable<TranslationEntry> ExtractStrings(ExtractableFolder strings)
    {
        var stringsRoot = strings.FullPath;
        RequiredMods? requiredMods = strings.RequiredPackageId == null
            ? null : new RequiredMods().Tap(rm => rm.AddAllowedByPackageIds(strings.RequiredPackageId.Split(',')));

        foreach (var filePath in FileInterface.DescendantFiles(stringsRoot).Where(x => x.ToLower().EndsWith(".txt")))
        {
            var nodeName = Path.GetRelativePath(stringsRoot, filePath);
            nodeName = Path.GetFileNameWithoutExtension(nodeName.Replace('\\', '.'));
            var lines = File.ReadAllLines(filePath);
            for (var i = 0; i < lines.Length; i++)
            {
                yield return new TranslationEntry("Strings", $"{nodeName}.{i}", lines[i], null, requiredMods, null);
            }
        }
    }

    private IEnumerable<TranslationEntry> ExtractPatches(SimulationResult simResult, ExtractableFolder patches, bool isOfficialContent)
    {
        simResult.DefsAddedByPatches.Clear();
        var patchesRoot = patches.FullPath;
        RequiredMods? requiredMods = patches.RequiredPackageId == null
            ? null : new RequiredMods().Tap(rm => rm.AddAllowedByPackageIds(patches.RequiredPackageId.Split(',')));

        var doc = new XDocument(new XElement("Patch"));
        foreach (var filePath in FileInterface.DescendantFiles(patchesRoot).Where(x => x.ToLower().EndsWith(".xml")))
        {
            var childDoc = FileInterface.ReadXml(filePath);
            foreach (var node in childDoc.Root!.Elements())
            {
                if (node.Name.LocalName == "Operation")
                    doc.Root!.Add(new XElement(node));
            }
        }

        foreach (var node in doc.Root!.Elements())
        {
            foreach (var translationEntry in PatchOperations.PatchOperationRecursive(node, simResult, null, false))
                yield return translationEntry with { RequiredMods = translationEntry.RequiredMods + requiredMods };
        }

        if (simResult.DefsAddedByPatches.Count == 0) yield break;

        foreach (var (requiredModsPatches, node) in simResult.DefsAddedByPatches)
        {
            if (requiredModsPatches != null)
                node.AppendElement("REQUIREDMODS", requiredModsPatches.ToString());

            if (node.Attribute("Name")?.Value is string name)
                simResult.ParentNodeLookUp[name] = node;
        }

        SimulatorEngine.DoXmlInheritance(simResult, simResult.DefsAddedByPatches.Select(x => x.Item2));

        foreach (var translation in ExtractDefs(simResult, isOfficialContent))
        {
            yield return translation with
            {
                ClassName = $"Patches.{translation.ClassName}",
                RequiredMods = translation.RequiredMods + requiredMods
            };
        }
    }
}