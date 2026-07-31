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
            var finalEntries = new List<TranslationEntry>();
            
            if (simResult.Snapshots.Count == 0)
                return new ExtractionResult(simResult.TargetMod, finalEntries);

            // Base 우주(인덱스 0) 추출
            var baseSnapshot = simResult.Snapshots[0];
            var baseEntries = ExtractionProcedureInjector.Instance.Execute(baseSnapshot, simResult.TargetMod)?.ToList() ?? new List<TranslationEntry>();
            
            // 어셈블리에서 Keyed 데이터를 직접 뽑아와서 베이스 리스트에 합침
            baseEntries.AddRange(KeyExtractor.Extract(baseSnapshot, simResult.TargetMod));
            
            // Base 항목들은 기본 리스트에 추가
            finalEntries.AddRange(baseEntries);

            // Base의 키와 원본 텍스트를 캐싱 (차분 비교용)
            var baseDictionary = new Dictionary<string, string>(); 
            foreach (var entry in baseEntries)
            {
                baseDictionary[entry.ClassNode] = entry.Original;
            }

            // 평행 우주 차분(Diff) 추출
            for (int i = 1; i < simResult.Snapshots.Count; i++)
            {
                var branchSnapshot = simResult.Snapshots[i];
                var branchEntries = ExtractionProcedureInjector.Instance.Execute(branchSnapshot, simResult.TargetMod)?.ToList() ?? new List<TranslationEntry>();

                // 분기 우주의 어셈블리 Keyed 데이터도 뽑아옵니다.
                branchEntries.AddRange(KeyExtractor.Extract(branchSnapshot, simResult.TargetMod));
                
                var branchCondition = new RequiredMods();
                branchCondition.AddAllowedByModNames(branchSnapshot.RequiredModIds);

                foreach (var branchEntry in branchEntries)
                {
                    bool isUniqueToBranch = false;

                    // 차분 조건 1: Base에 아예 없는 새로운 번역 키인가?
                    if (!baseDictionary.TryGetValue(branchEntry.ClassNode, out var baseOriginalText))
                    {
                        isUniqueToBranch = true;
                    }
                    // 차분 조건 2: 키는 있는데 원본 텍스트가 Base와 다른가?
                    else if (baseOriginalText != branchEntry.Original)
                    {
                        isUniqueToBranch = true;
                    }

                    if (isUniqueToBranch)
                    {
                        // 이 분기에만 존재하는 고유한 번역이므로 조건 꼬리표를 달아서 추가
                        finalEntries.Add(branchEntry with { RequiredMods = branchCondition });
                    }
                }
            }

            // 중복 필터링 및 후처리 파이프라인
            var distinctEntries = FilterDuplicates(finalEntries);
            var processedEntries = TranslationEntryProcedureInjector.Instance.Execute(distinctEntries).ToList();

            return new ExtractionResult(simResult.TargetMod, processedEntries);
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

        public static IEnumerable<TranslationEntry> ExtractKeyed(ExtractableFolder keyed, bool isOfficialContent)
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

        public static IEnumerable<TranslationEntry> ExtractStrings(ExtractableFolder strings)
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