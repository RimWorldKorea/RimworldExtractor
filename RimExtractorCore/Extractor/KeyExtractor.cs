using ICSharpCode.Decompiler;
using ICSharpCode.Decompiler.Metadata;
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
        public static IEnumerable<TranslationEntry> Extract(DefSnapshot snapshot, Mod targetMod)
        {
            var extractedKeys = new Dictionary<string, TranslationEntry>();
            var requiredMods = new RequiredMods();
            requiredMods.AddAllowedByModNames(snapshot.RequiredModIds);
            
            Log.Msg($"KeyExtractor.Extract -> {requiredMods.ToString()}");
            
            var assemblyPaths = new List<string>();

            // 1. DLL 경로 수집
            foreach (var unit in snapshot.LoadedUnits)
            {
                if (unit.Assemblies != null && unit.Assemblies.Exists)
                    assemblyPaths.AddRange(unit.Assemblies.GetFiles("*.dll", SearchOption.AllDirectories).Select(f => f.FullName));
            }

            foreach (var dllPath in assemblyPaths.Distinct(StringComparer.OrdinalIgnoreCase))
                ExtractKeysFromAssembly(dllPath, extractedKeys, requiredMods);

            // 3. 설정된 1차/2차 언어를 기반으로 원문(Original) 텍스트 매핑
            var priorityLanguages = SettingManager.Current.GetLanguagePriorityList().ToList();
            
            foreach (var unit in snapshot.LoadedUnits)
            {
                foreach (var lang in priorityLanguages)
                {
                    var englishName = lang.Split('(')[0].Trim();
                    if(!LanguageInfo.TryParse(englishName, out var languageInfo)) continue;
                    
                    var keyedDir = unit.GetKeyedPath(languageInfo);
                    if (keyedDir == null || !keyedDir.Exists) continue;

                    foreach (var xmlFile in keyedDir.GetFiles("*.xml", SearchOption.AllDirectories))
                    {
                        var doc = FileInterface.ReadXml(xmlFile.FullName);
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
                                    extractedKeys[key] = entry with
                                    {
                                        Original = node.Value, SourceFile =
                                        Path.GetFileNameWithoutExtension(xmlFile.FullName)
                                    };
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
            try
            {
                // 1. 대상 어셈블리를 PE 파일 모듈로 로드
                using var module = new PEFile(dllPath);

                // [수정됨] module.Reader가 아니라 module 객체에서 직접 호출합니다. 
                // (만약 그래도 빨간 줄이 뜬다면, 그냥 null을 전달하셔도 림월드 환경에서는 무방합니다.)
                var targetFramework = module.DetectTargetFrameworkId();
                var resolver = new UniversalAssemblyResolver(dllPath, throwOnError: false, targetFramework);
                
                // 3. 리졸버의 탐색 경로에 림월드 본편 Managed 폴더(Assembly-CSharp.dll 위치) 강제 추가
                var managedDir = Path.Combine(SettingManager.Current.PathRimworld, "RimWorldWin64_Data", "Managed");
                if (Directory.Exists(managedDir))
                {
                    resolver.AddSearchDirectory(managedDir);
                }

                // 4. 모듈과 커스텀 리졸버를 주입하여 디컴파일러 생성
                var decompiler = new CSharpDecompiler(module, resolver, new DecompilerSettings(LanguageVersion.Latest));
                var syntaxTree = decompiler.DecompileWholeModuleAsSingleFile();

                // 5. AST 분석 시작
                foreach (var invocation in syntaxTree.Descendants.OfType<InvocationExpression>())
                {
                    if (invocation.Target is MemberReferenceExpression memberRef)
                    {
                        var methodName = memberRef.MemberName;
                        if (methodName == "Translate" || methodName == "TryTranslate" || methodName == "TranslateWithBackup" || methodName == "TranslateSimple")
                        {
                            // Case 1. "MyKey".Translate() 처럼 대상이 순수한 문자열(PrimitiveExpression)인 경우
                            if (memberRef.Target is PrimitiveExpression primitiveTarget && primitiveTarget.Value is string targetKey)
                            {
                                AddKey(targetKey, dict, requiredMods);
                            }
                            // Case 2. 혹시 모를 Translate("MyKey") 형태 대응 (인자가 순수 문자열인 경우)
                            else if (invocation.Arguments.FirstOrDefault() is PrimitiveExpression primitiveArg && primitiveArg.Value is string argKey)
                            {
                                AddKey(argKey, dict, requiredMods);
                            }
                            
                            // 그 외의 경우 (변수 사용, switch 구문, 문자열 보간 등)는 
                            // 추출이 불가능하거나 런타임 결정 값이므로 로그 없이 조용히 무시(Pass)합니다.
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Log.Err($"[Keyed 추출] 어셈블리({Path.GetFileName(dllPath)}) 디컴파일 중 예외 발생: {e.Message}");
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