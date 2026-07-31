using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ICSharpCode.Decompiler;
using ICSharpCode.Decompiler.CSharp;
using ICSharpCode.Decompiler.CSharp.Syntax;
using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.Extractor
{
    /// <summary>
    /// 모드 어셈블리(dll)의 IL 코드를 분석하여 하드코딩된 Keyed 번역 키를 정적으로 추출하는 내부 엔진입니다.
    /// </summary>
    public static class KeyExtractor
    {
        public static IEnumerable<TranslationEntry> Extract(DefSnapshot snapshot, ModMetadata targetMod)
        {
            var extractedKeys = new Dictionary<string, TranslationEntry>();
            var requiredMods = new RequiredMods();
            requiredMods.AddAllowedByModNames(snapshot.RequiredModIds);

            // 1. 현재 우주(Snapshot)에 할당된 폴더들에서 Assemblies 탐색
            var assemblyPaths = new List<string>();
            foreach (var folder in snapshot.AssignedFolders)
            {
                var asmDir = Path.Combine(folder.FullPath, "Assemblies");
                if (Directory.Exists(asmDir))
                {
                    assemblyPaths.AddRange(Directory.GetFiles(asmDir, "*.dll", SearchOption.AllDirectories));
                }
            }

            // 2. 어셈블리 디컴파일 및 AST 순회
            foreach (var dllPath in assemblyPaths)
            {
                ExtractKeysFromAssembly(dllPath, extractedKeys, requiredMods);
            }

            // 3. 설정된 1차/2차 언어를 기반으로 원문(Original) 텍스트 매핑
            var priorityLanguages = SettingManager.Current.GetLanguagePriorityList().ToList();
            
            foreach (var folder in snapshot.AssignedFolders)
            {
                foreach (var lang in priorityLanguages)
                {
                    var keyedDir = Path.Combine(folder.FullPath, "Languages", lang, "Keyed");
                    if (!Directory.Exists(keyedDir)) continue;

                    foreach (var xmlPath in Directory.GetFiles(keyedDir, "*.xml", SearchOption.AllDirectories))
                    {
                        var doc = FileInterface.ReadXml(xmlPath);
                        if (doc.Root == null) continue;

                        foreach (var node in doc.Root.Elements())
                        {
                            var key = node.Name.LocalName;
                            
                            // 어셈블리에서 발견된 키인 경우
                            if (extractedKeys.TryGetValue(key, out var entry))
                            {
                                // 원문이 비어있다면 우선순위가 높은 언어의 텍스트로 채움
                                if (string.IsNullOrEmpty(entry.Original))
                                {
                                    extractedKeys[key] = entry with { Original = node.Value, SourceFile = Path.GetFileNameWithoutExtension(xmlPath) };
                                }
                            }
                        }
                    }
                }
            }

            return extractedKeys.Values;
        }

        private static void ExtractKeysFromAssembly(string dllPath, Dictionary<string, TranslationEntry> dict, RequiredMods requiredMods)
        {
            Log.Msg($"[Keyed 추출] 어셈블리 분석 중: {Path.GetFileName(dllPath)}");
            var decompiler = new CSharpDecompiler(dllPath, new DecompilerSettings(LanguageVersion.Latest));
            var syntaxTree = decompiler.DecompileWholeModuleAsSingleFile();

            foreach (var invocation in syntaxTree.Descendants.OfType<InvocationExpression>())
            {
                if (invocation.Target is MemberReferenceExpression memberRef)
                {
                    var methodName = memberRef.MemberName;

                    if (methodName == "Translate" || methodName == "TryTranslate" || methodName == "TranslateWithBackup" || methodName == "TranslateSimple")
                    {
                        if (memberRef.Target is PrimitiveExpression primitiveTarget && primitiveTarget.Value is string targetKey)
                        {
                            AddKey(targetKey, dict, requiredMods);
                        }
                        else if (invocation.Arguments.FirstOrDefault() is PrimitiveExpression primitiveArg && primitiveArg.Value is string argKey)
                        {
                            AddKey(argKey, dict, requiredMods);
                        }
                        else
                        {
                            Log.Wrn($"동적으로 생성되는 번역 키(Keyed)가 감지되었습니다. 추출에서 제외됩니다: {invocation}");
                        }
                    }
                }
            }
        }

        private static void AddKey(string key, Dictionary<string, TranslationEntry> dict, RequiredMods requiredMods)
        {
            if (!dict.ContainsKey(key))
            {
                dict[key] = new TranslationEntry("Keyed", key, "", null, requiredMods, null);
            }
        }
    }
}