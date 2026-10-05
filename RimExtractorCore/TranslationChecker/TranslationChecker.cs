using System.Text.RegularExpressions;

namespace RimExtractorCore.TranslationChecker;

/// <summary>번역 결과물을 검증하는 간단한 도구입니다.</summary>
public static class TranslationChecker
{
    /// <summary>번역 파일을 읽고 기본적인 확인 사항을 점검합니다. 결과는 로그로 출력합니다.</summary>
    /// <param name="pathContainsLanguages">Languages 폴더가 위치한 경로입니다.</param>
    public static void CheckTranslation(string pathContainsLanguages)
    {
        Log.Msg("번역 기본 검증 시작");

        // 1. LoadFolders.Build.yaml 기초 정보 확인
        CheckLoadFoldersBuildYaml(pathContainsLanguages);

        // 2. Languages 폴더 및 하위 언어 폴더 확인
        var languagesDir = Path.Combine(pathContainsLanguages, "Languages");
        if (!Directory.Exists(languagesDir))
        {
            Log.Err($"Languages 폴더를 찾을 수 없습니다.");
            Log.Good("번역 기본 검증 완료");
            return;
        }

        var foundLangFolders = CheckLanguageFolders(languagesDir);

        // 3. 각 언어 폴더별 XML 파일 텍스트 정밀 검증
        foreach (var langFolder in foundLangFolders)
        {
            CheckXmlFiles(langFolder);
        }

        Log.Good("번역 기본 검증 완료");
    }
    
    private static void CheckLoadFoldersBuildYaml(string rootPath)
    {
        var yamlPath = Path.Combine(rootPath, "LoadFolders.Build.yaml");
        if (File.Exists(yamlPath))
        {
            Log.Good("LoadFolders.Build.yaml 파일 확인.");
            
            var lines = File.ReadAllLines(yamlPath);
            string packageId = "알 수 없음";
            string modName = "알 수 없음";

            foreach (var line in lines)
            {
                // 정규식을 통해 값만 빠르게 추출
                if (line.Contains("PackageID:"))
                {
                    var match = Regex.Match(line, @"PackageID:\s*\[""(.*?)""\]");
                    if (match.Success) packageId = match.Groups[1].Value;
                }
                else if (line.Contains("ModName:"))
                {
                    var match = Regex.Match(line, @"ModName:\s*""(.*?)""");
                    if (match.Success) modName = match.Groups[1].Value;
                }
            }

            Log.Msg($" 🔹 모드 이름 : {modName}");
            Log.Msg($" 🔹 Package ID : {packageId}");
        }
        else
        {
            Log.Wrn("LoadFolders.Build.yaml 파일이 없습니다.\nRMK에 새로운 번역을 추가하고자 하는 경우 빌드 파일을 작성해주시기 바랍니다.");
        }
    }
    
    private static List<string> CheckLanguageFolders(string languagesDir)
    {
        var foundPaths = new List<string>();
        var dirs = Directory.GetDirectories(languagesDir);

        if (dirs.Length == 0)
        {
            Log.Wrn("⚠️ Languages 폴더 내에 번역된 하위 폴더(예: Korean (한국어))가 존재하지 않습니다.");
            return foundPaths;
        }

        Log.Msg("✅ 다음 언어에 대한 번역 폴더가 발견되었습니다:");
        foreach (var dir in dirs)
        {
            var dirName = Path.GetFileName(dir);
            
            // 등록된 표준 언어 목록과 대조 (Korean, Korean (한국어) 등 호환)
            var knownLang = LanguageInfo.SupportedLanguages.FirstOrDefault(l =>
                string.Equals(l.EnglishName, dirName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(l.ExtendedName, dirName, StringComparison.OrdinalIgnoreCase));

            if (knownLang != null)
            {
                Log.Msg($" 🔹 {knownLang.ExtendedName} (올바른 표준 폴더명)");
            }
            else
            {
                Log.Wrn($" ⚠️ {dirName} (표준 언어 폴더명이 아닐 수 있습니다. 게임에서 인식하지 못할 수 있으니 폴더명을 확인하세요.)");
            }
            
            foundPaths.Add(dir);
        }

        return foundPaths;
    }

    private static void CheckXmlFiles(string langFolderPath)
    {
        var xmlFiles = Directory.GetFiles(langFolderPath, "*.xml", SearchOption.AllDirectories);
        if (xmlFiles.Length == 0) return;

        Log.Msg($"\n🔍 [{Path.GetFileName(langFolderPath)}] 언어 파일 텍스트 정밀 검사를 시작합니다...");

        int warningCount = 0;

        // XML 태그 추출을 위한 정규식 (예: <label>텍스트</label>)
        // (태그 내부 속성 포함, 멀티라인이 아닌 일반적인 단일 줄 번역 위주 탐지)
        var xmlTagRegex = new Regex(@"<([a-zA-Z0-9_\-\.]+)(?:\s+[^>]*)?>(.*?)</\1>");

        // 1. 앰퍼샌드 오류 점검 (& 기호가 단독으로 쓰인 경우 파싱 에러 발생)
        var unescapedAmpersandRegex = new Regex(@"&(?!amp;|lt;|gt;|quot;|apos;)");

        // 2. 림월드 한국어 잘못된 조사 형식 정규식 (을(를), 이(가), 은(는), 와(과))
        var badParticleRegex = new Regex(@"을\(를\)|이\(가\)|은\(는\)|와\(과\)");

        foreach (var file in xmlFiles)
        {
            var lines = File.ReadAllLines(file);
            var relativePath = Path.GetRelativePath(langFolderPath, file);

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var match = xmlTagRegex.Match(line);
                
                if (match.Success)
                {
                    var tag = match.Groups[1].Value;
                    var text = match.Groups[2].Value; // 디코딩/트림이 전혀 발생하지 않은 원시 텍스트

                    if (string.IsNullOrEmpty(text)) continue;

                    // [체크 1] 앞뒤 화이트스페이스 검사
                    if (text.StartsWith(" ") || text.EndsWith(" ") || text.StartsWith("\t") || text.EndsWith("\t"))
                    {
                        Log.Wrn($"[{relativePath} : {i + 1}줄] <{tag}> 텍스트 앞뒤에 불필요한 공백이나 탭이 있습니다.");
                        warningCount++;
                    }

                    // [체크 2] 이스케이프되지 않은 예약어 '&' 검사
                    if (unescapedAmpersandRegex.IsMatch(text))
                    {
                        Log.Wrn($"[{relativePath} : {i + 1}줄] <{tag}> 텍스트에 처리되지 않은 '&' 예약어가 있습니다. 게임 오류를 방지하려면 '&amp;'로 치환해야 합니다.");
                        warningCount++;
                    }

                    // [체크 3] 한국어 조사 검사 (한국어 폴더인 경우만 작동하도록 해도 되지만, 텍스트 자체를 스캔하므로 공통 적용)
                    if (badParticleRegex.IsMatch(text))
                    {
                        Log.Wrn($"[{relativePath} : {i + 1}줄] <{tag}> 텍스트에 '을(를)'과 같은 잘못된 조사 형식이 있습니다. 림월드 공식 표준인 '(을)를', '(이)가' 형식으로 수정을 권장합니다.");
                        warningCount++;
                    }
                }
            }
        }

        if (warningCount == 0)
        {
            Log.Msg($"✅ [{Path.GetFileName(langFolderPath)}] 텍스트 검증 완료. 의심되는 문제점을 발견하지 못했습니다.");
        }
        else
        {
            Log.Wrn($"⚠️ [{Path.GetFileName(langFolderPath)}] 텍스트 검증 완료. 총 {warningCount}건의 확인 사항이 발견되었습니다. 로그를 올려서 확인해 보세요!");
        }
    }
}