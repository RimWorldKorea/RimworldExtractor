// Extractor/Extractor.cs 

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;
using RimExtractorCore.Procedures;

namespace RimExtractorCore.Extractor
{
    public static partial class ExtractorEngine
    {
        public static ExtractionResult ExtractTranslationData(SimulationResult simResult)
        {
            var finalEntries = new List<TranslationEntry>();

            if (simResult.Snapshots.Count == 0)
                return new ExtractionResult(simResult.TargetMod, finalEntries);

            // 1. Base snapshot (인덱스 0) 추출
            var baseSnapshot = simResult.Snapshots[0];
            var baseEntries = ExtractFromSnapshot(baseSnapshot, simResult.TargetMod);

            finalEntries.AddRange(baseEntries);

            var baseDictionary = new Dictionary<string, string>();
            foreach (var entry in baseEntries)
            {
                baseDictionary[entry.ClassNode] = entry.Original;
            }

            // 2. 다중 우주 분기(Diff) 처리 - 조건부 패치로 인해 변경/추가된 번역만 수집
            for (int i = 1; i < simResult.Snapshots.Count; i++)
            {
                var branchSnapshot = simResult.Snapshots[i];
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

        private static List<TranslationEntry> ExtractFromSnapshot(DefSnapshot snapshot, ModMetadata targetMod)
        {
            var entries = new List<TranslationEntry>();
            bool isOfficialContent = targetMod.IsOfficialContent;

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
        private static void ProcessLanguageData(List<TranslationEntry> entries, DefSnapshot snapshot, bool isOfficialContent)
        {
            var priorityLanguages = SettingManager.Current.GetLanguagePriorityList().ToList();
            
            // 빠른 검색 및 수정을 위한 딕셔너리 구성 (Key: ClassNode)
            var entryMap = new Dictionary<string, TranslationEntry>();
            foreach (var entry in entries)
            {
                entryMap[entry.ClassNode] = entry;
            }

            // snapshot에 포함된 실제 동작 버전의 루트 디렉토리들 추출 (예: C:\Mod\1.5)
            var versionDirs = snapshot.AssignedFolders
                .Select(f => 
                {
                    var path = f.FullPath;
                    var langIdx = path.IndexOf($"{Path.DirectorySeparatorChar}Languages{Path.DirectorySeparatorChar}");
                    if (langIdx >= 0) return path.Substring(0, langIdx);
                    return Path.GetDirectoryName(path);
                })
                .Where(d => !string.IsNullOrEmpty(d))
                .Distinct()
                .ToList();

            // Language 우선순위가 높은 것부터 채우기
            foreach (var lang in priorityLanguages)
            {
                var shortLang = lang.Split(' ').First(); // "Korean (한국어)" -> "Korean"
                var langNames = new HashSet<string> { lang, shortLang };

                foreach (var versionDir in versionDirs)
                {
                    foreach (var langName in langNames)
                    {
                        var langDir = Path.Combine(versionDir!, "Languages", langName);
                        if (!Directory.Exists(langDir)) continue;

                        // 1. DefInjected 읽기 (비어있는 Def 원문 채우기)
                        var defInjectedDir = Path.Combine(langDir, "DefInjected");
                        if (Directory.Exists(defInjectedDir))
                        {
                            foreach (var xmlPath in FileInterface.DescendantFiles(defInjectedDir).Where(x => x.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
                            {
                                try
                                {
                                    var className = Path.GetRelativePath(defInjectedDir, xmlPath).Split(Path.DirectorySeparatorChar).First();
                                    var doc = FileInterface.ReadXml(xmlPath);
                                    var parsed = LanguageXmlProcessor.ParseDefInjected(doc, className);
                                    
                                    foreach (var p in parsed)
                                    {
                                        if (entryMap.TryGetValue(p.ClassNode, out var existingEntry))
                                        {
                                            if (string.IsNullOrWhiteSpace(existingEntry.Original))
                                            {
                                                var text = p.Translated ?? p.Original;
                                                if (!string.IsNullOrEmpty(text))
                                                {
                                                    entryMap[p.ClassNode] = existingEntry with { Original = text };
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception e) { Log.Wrn($"DefInjected 분석 오류 ({xmlPath}): {e.Message}"); }
                            }
                        }

                        // 2. Keyed 읽기 (비어있는 IL Keyed 원문 채우기 + XML에만 있는 Keyed 추가)
                        var keyedDir = Path.Combine(langDir, "Keyed");
                        if (Directory.Exists(keyedDir))
                        {
                            var requiredMods = snapshot.AssignedFolders.FirstOrDefault(x => x.FullPath == keyedDir)?.RequiredPackageId;
                            var rm = requiredMods != null ? new RequiredMods().Tap(r => r.AddAllowedByPackageIds(requiredMods.Split(','))) : null;
                            
                            var keyedEntries = LanguageXmlProcessor.ParseKeyed(keyedDir, rm, isOfficialContent);
                            foreach (var kEntry in keyedEntries)
                            {
                                if (entryMap.TryGetValue(kEntry.ClassNode, out var existingEntry))
                                {
                                    // 기존 IL 추출 Keyed의 빈 원문을 덮어씀
                                    if (string.IsNullOrWhiteSpace(existingEntry.Original) && !string.IsNullOrEmpty(kEntry.Original))
                                    {
                                        entryMap[kEntry.ClassNode] = existingEntry with { Original = kEntry.Original, SourceFile = kEntry.SourceFile ?? existingEntry.SourceFile };
                                    }
                                }
                                else
                                {
                                    // XML에만 존재하는 Keyed 항목 추가
                                    entryMap[kEntry.ClassNode] = kEntry;
                                }
                            }
                        }

                        // 3. Strings 읽기 (TXT 파일에서 추출 및 추가)
                        var stringsDir = Path.Combine(langDir, "Strings");
                        if (Directory.Exists(stringsDir))
                        {
                            var requiredMods = snapshot.AssignedFolders.FirstOrDefault(x => x.FullPath == stringsDir)?.RequiredPackageId;
                            var rm = requiredMods != null ? new RequiredMods().Tap(r => r.AddAllowedByPackageIds(requiredMods.Split(','))) : null;

                            var stringsEntries = LanguageXmlProcessor.ParseStrings(stringsDir, rm);
                            foreach (var sEntry in stringsEntries)
                            {
                                if (!entryMap.ContainsKey(sEntry.ClassNode))
                                {
                                    entryMap[sEntry.ClassNode] = sEntry;
                                }
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
            var set = new HashSet<(string, string)>();
            var distinctList = new List<TranslationEntry>();
            foreach (var entry in extraction)
            {
                var tuple = (entry.ClassNode, entry.Original);
                var pair = set.FirstOrDefault(x => x.Item1 == tuple.Item1);
                
                if (pair != default)
                {
                    if (pair.Item2 != entry.Original)
                        Log.Err($"중복 노드 감지: {entry.ClassNode} | 기존 원문: {pair.Item2} | 새 원문: {entry.Original}");
                }
                else
                {
                    set.Add(tuple);
                    distinctList.Add(entry);
                }
            }
            return distinctList;
        }
    }
}