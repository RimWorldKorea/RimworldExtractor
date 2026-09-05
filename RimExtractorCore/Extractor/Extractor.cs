using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.Extractor;

public static partial class ExtractorEngine
{
    public static ExtractionResult ExtractTranslationData(SimulationResult simResult)
    {
        var finalEntries = new List<TranslationEntry>();

        if (simResult.Snapshots.Count == 0)
            return new ExtractionResult(simResult.TargetMod, finalEntries);

        // Base snapshot을 플래그로 명확하게 탐색 (인덱스 의존 탈피)
        var baseSnapshot = simResult.Snapshots.FirstOrDefault(s => s.IsBaseSnapshot);
        //TODO 베이스 스냅샷이 없을 때 이렇게 처리하는게 맞을까?
        if (baseSnapshot == null)
        {
            Log.Err("베이스 스냅샷을 찾을 수 없습니다!");
            return new ExtractionResult(simResult.TargetMod, finalEntries);
        }

        Log.Common($" 빙빙빙");
        var baseEntries = ExtractFromSnapshot(baseSnapshot, simResult.TargetMod);

        finalEntries.AddRange(baseEntries);

        var baseDictionary = new Dictionary<string, string>();
        foreach (var entry in baseEntries)
        {
            baseDictionary[entry.ClassNode] = entry.Original;
        }

        // 2. 다중 우주 분기(Diff) 처리 - 조건부 패치로 인해 변경/추가된 번역만 수집
        foreach (var branchSnapshot in simResult.Snapshots.Where(s => !s.IsBaseSnapshot))
        {
            var branchEntries = ExtractFromSnapshot(branchSnapshot, simResult.TargetMod);

            var branchCondition = new RequiredMods();
            branchCondition.AddAllowedByModNames(branchSnapshot.RequiredModIds);

            foreach (var branchEntry in branchEntries)
            {
                bool isUniqueToBranch = false;

                if (!baseDictionary.TryGetValue(branchEntry.ClassNode, out var baseOriginalText))
                {
                    isUniqueToBranch = true;
                }
                else if (baseOriginalText != branchEntry.Original)
                {
                    isUniqueToBranch = true;
                }

                if (isUniqueToBranch)
                {
                    finalEntries.Add(branchEntry with { RequiredMods = branchCondition });
                }
            }
        }

        // 3. 중복 제거 및 최종 프로시저(Stage C) 필터링
        var distinctEntries = FilterDuplicates(finalEntries);
        var processedEntries = TranslationEntryProcedureInjector.Instance.Execute(distinctEntries).ToList();

        return new ExtractionResult(simResult.TargetMod, processedEntries);
    }

    private static List<TranslationEntry> ExtractFromSnapshot(DefSnapshot snapshot, Mod targetMod)
    {
        var entries = new List<TranslationEntry>();
        bool isOfficialContent = targetMod.Type == ModType.Official;
        
        // Step 1: Defs 추출 (스키마 기반 DefaultExtractionProcedure)
        var defEntries = ExtractionProcedureInjector.Instance.Execute(snapshot, targetMod);
        if (defEntries != null) entries.AddRange(defEntries);

        // Step 2: Keyed 추출 (IL 어셈블리 분석 - 원문이 비어있음)
        entries.AddRange(KeyExtractor.Extract(snapshot, targetMod));

        // Step 3: LanguageData 폴더를 순회하며 빈 원문을 채우고 XML 기반 Keyed, Strings를 추가
        ProcessLanguageData(entries, snapshot, isOfficialContent);

        return entries;
    }

    /// <summary>
    /// 추출된 항목들의 비어있는 Original 값을 채우고, 누락된 XML Keyed와 Strings를 추가합니다.
    /// </summary>
    private static void ProcessLanguageData(List<TranslationEntry> entries, DefSnapshot snapshot,
        bool isOfficialContent)
    {
        var priorityLanguages = SettingManager.Current.GetLanguagePriorityList().ToList();

        // 빠른 검색 및 수정을 위한 딕셔너리 구성 (Key: ClassNode)
        var entryMap = new Dictionary<string, TranslationEntry>();
        foreach (var entry in entries) entryMap[entry.ClassNode] = entry;
        
        // Language 우선순위가 높은 것부터 채우기
        foreach (var lang in priorityLanguages)
        {
            var englishName = lang.Split('(')[0].Trim();
            var languageInfo = LanguageInfo.Get(englishName);

            foreach (var unit in snapshot.LoadedUnits)
            {
                // 1. Keyed XML 병합
                var keyedDir = unit.GetKeyedPath(languageInfo);
                if (keyedDir != null && keyedDir.Exists)
                {
                    var rm = new RequiredMods();
                    rm.AddAllowedByModNames(snapshot.RequiredModIds);

                    var keyedEntries = LanguageXmlProcessor.ParseKeyed(keyedDir.FullName, rm, isOfficialContent);
                    foreach (var kEntry in keyedEntries)
                    {
                        if (entryMap.TryGetValue(kEntry.ClassNode, out var existingEntry))
                        {
                            if (string.IsNullOrWhiteSpace(existingEntry.Original) && !string.IsNullOrEmpty(kEntry.Original))
                            {
                                entryMap[kEntry.ClassNode] = existingEntry with { Original = kEntry.Original, SourceFile = kEntry.SourceFile ?? existingEntry.SourceFile };
                            }
                        }
                        else { entryMap[kEntry.ClassNode] = kEntry; }
                    }
                }

                // 2. Strings TXT 병합
                var localeDir = unit.GetLocaleFolder(languageInfo);
                if (localeDir != null && localeDir.Exists)
                {
                    var stringsDir = new DirectoryInfo(Path.Combine(localeDir.FullName, "Strings"));
                    if (stringsDir.Exists)
                    {
                        var rm = new RequiredMods();
                        rm.AddAllowedByModNames(snapshot.RequiredModIds);

                        var stringsEntries = LanguageXmlProcessor.ParseStrings(stringsDir.FullName, rm);
                        foreach (var sEntry in stringsEntries)
                        {
                            if (!entryMap.ContainsKey(sEntry.ClassNode)) entryMap[sEntry.ClassNode] = sEntry;
                        }
                    }
                }
            }
        }

        // 결과를 다시 리스트에 덮어쓰기
        entries.Clear();
        entries.AddRange(entryMap.Values);
    }

    private static List<TranslationEntry> FilterDuplicates(List<TranslationEntry> extraction)
    {
        var distinctList = new List<TranslationEntry>();
        
        // 고유 식별 키: (노드 경로, 원본 텍스트, 요구 모드 문자열)
        var handledSet = new HashSet<(string ClassNode, string Original, string ReqMods)>();

        foreach (var entry in extraction)
        {
            var reqModsStr = entry.RequiredMods?.ToString() ?? string.Empty;
            var key = (entry.ClassNode, entry.Original, reqModsStr);

            // 동일한 노드 경로에 대해 원문 텍스트나 요구 모드가 하나라도 다르면 새로운 엔트리로 인정합니다.
            if (!handledSet.Contains(key))
            {
                handledSet.Add(key);
                distinctList.Add(entry);
            }
        }

        return distinctList;
    }
}