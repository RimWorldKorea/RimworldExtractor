using System.IO;
using System.Linq;
using System.Collections.Generic;
using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;
using RimExtractorCore.Extractor;

namespace RimExtractorCore.Procedures
{
    public class DefaultExtractionProcedure : IExtractionProcedure
    {
        public string Name => "DefaultExtractionProcedure";

        public IEnumerable<TranslationEntry> Extract(DefSnapshot snapshot, ModMetadata targetMod)
        {
            var extraction = new List<TranslationEntry>();
            bool isOfficialContent = targetMod.IsOfficialContent;

            // 1. Defs 추출 (이미 패치와 상속이 모두 끝난 순수한 트리)
            extraction.AddRange(ExtractDefs(snapshot, isOfficialContent));

            // 2. Keyed, Strings 추출 (이 우주에 할당된 폴더만 순회)
            foreach (var extractableFolder in snapshot.AssignedFolders)
            {
                string folderName = Path.GetFileName(extractableFolder.FolderName);
                if (folderName == "Keyed")
                {
                    extraction.AddRange(ExtractorEngine.ExtractKeyed(extractableFolder, isOfficialContent));
                }
                else if (folderName == "Strings")
                {
                    extraction.AddRange(ExtractorEngine.ExtractStrings(extractableFolder));
                }
            }

            return extraction;
        }

        internal IEnumerable<TranslationEntry> ExtractDefs(DefSnapshot snapshot, bool isOfficialContent)
        {
            if (snapshot.Tree.Root == null) yield break;

            foreach (var node in snapshot.Tree.Root.Elements().Where(x => x.Attribute("Reference")?.Value.ToLower() != "true"))
            {
                var defName = node.Element("defName")?.Value;
                if (defName == null)
                {
                    if (node.Name.LocalName != "SongDef")
                        Log.Wrn($"SongDef가 아닌 Abstract가 아닌 XML 노드에 'defName'이 없습니다. InnerXml: {node}");
                    continue;
                }

                var className = node.Attribute("Class")?.Value ?? node.Name.LocalName;
                className = className[..1].ToUpper() + className[1..];

                // ExtractorEngine (label, description 등 텍스트 추출)
                foreach (var translationEntry in ExtractorEngine.FindExtractableNodes(defName, className, node, isOfficialContent))
                {
                    yield return translationEntry;
                }
            }
        }
    }
}