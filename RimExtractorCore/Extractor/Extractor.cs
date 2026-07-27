using RimExtractorCore.DataTypes;

namespace RimExtractorCore.Extractor

{
    public static partial class ExtractorEngine
    {
        /// <summary>
        /// 시뮬레이션 결과(DefTree)를 받아 최종 가공된 TranslationEntry 컨테이너를 반환합니다.
        /// </summary>
        public static ExtractionResult ExtractTranslationData(SimulationResult simResult)
        {
            // 1. 단일 추출 프로시저 실행 (원시 데이터 추출)
            var rawEntries = ExtractionProcedureInjector.Instance.Execute(simResult)?.ToList();

            if (rawEntries == null)
            {
                Log.Err("PrimaryExtractor가 로드되지 않았습니다.");
                return new ExtractionResult(simResult.TargetMod!, new List<TranslationEntry>());
            }
            
            // 2. 고유성 검증 및 중복 필터링
            var distinctEntries = FilterDuplicates(rawEntries);

            // 3. 다중 후처리 파이프라인 실행 (원시 데이터 가공/필터링)
            // (MVCF, NodeReplacement 등의 ITranslationProcedure들이 통합적으로 1회 순차 실행됨)
            var processedEntries = TranslationEntryProcedureInjector.Instance.Execute(distinctEntries).ToList();

            // 4. 추출된 데이터를 컨테이너에 담아 반환 (파일 IO 역할 완벽히 분리)
            return new ExtractionResult(simResult.TargetMod!, processedEntries);
        }

        private static List<TranslationEntry> FilterDuplicates(List<TranslationEntry> extraction)
        {
            var set = new HashSet<(string, string)>();
            var distinctList = new List<TranslationEntry>();

            foreach (var entry in extraction)
            {
                var tuple = (entry.ClassName + "+" + entry.Node, entry.Original);
                var pair = set.FirstOrDefault(x => x.Item1 == tuple.Item1);

                if (pair != default)
                {
                    if (pair.Item2 != entry.Original)
                        Log.Err($"중복 키 발생 및 원문 불일치: {entry.ClassName}+{entry.Node}| 기존: {pair.Item2} | 새 원문: {entry.Original}");
                }
                else
                {
                    set.Add(tuple);
                    distinctList.Add(entry);
                }
            }
            return distinctList;
        }

        internal static IEnumerable<TranslationEntry> ExtractKeyed(ExtractableFolder keyed, bool isOfficialContent)
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

        internal static IEnumerable<TranslationEntry> ExtractStrings(ExtractableFolder strings)
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
        
        public static T Tap<T>(this T obj, Action<T> action) { action(obj); return obj; }
    }
}