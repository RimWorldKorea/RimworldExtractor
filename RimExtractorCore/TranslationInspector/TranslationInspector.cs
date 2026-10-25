using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace RimExtractorCore.TranslationInspector;

/// <summary>번역 결과물을 검증하는 간단한 도구입니다.</summary>
public static class TranslationInspector
{
    /// <summary>잘못된 예약어 검사 패턴</summary>
    private static readonly Regex UnescapedAmpersandRegex = new (
        @"&(?!amp;|lt;|gt;|quot;|apos;|#)", 
        RegexOptions.Compiled);

    /// <summary>잘못된 조사 검사 패턴</summary>
    /// <remarks>짝으로 된 조사 키워드의 순서나 괄호 위치를 검사합니다.</remarks>
    private static readonly Regex BadParticleFormatRegex = new (
        @"을\(를\)|\(를\)을|를\(을\)|이\(가\)|\(가\)이|가\(이\)|은\(는\)|\(는\)은|는\(은\)|와\(과\)|\(과\)와|과\(와\)|으\(로\)|이\(어\)", 
        RegexOptions.Compiled);

    /// <summary>잘못된 조사 검사 패턴 3</summary>
    /// <remarks>공백이 끼어 동작하지 않는 자동 조사를 검사합니다.</remarks>
    private static readonly Regex SpaceBeforeParticleRegex = new (
        @"[\}\]]\s+(\(을\)를|\(이\)가|\(은\)는|\(와\)과|\(으\)로)", 
        RegexOptions.Compiled);
    
    /// <summary>잘못된 조사 검사 패턴 2</summary>
    /// <remarks>짝으로 되지 않은 조사 키워드를 검사합니다.</remarks>
    private static readonly Regex SingleParticleAfterVarRegex = new (
        @"[\}\]]\s*(을|를|이|가|은|는|와|과|으로|로)(?!\w)", 
        RegexOptions.Compiled);
    
    /// <summary>번역 파일을 읽고 기본적인 확인 사항을 점검합니다. 결과는 로그로 출력합니다.</summary>
    /// <param name="pathContainsLanguages">Languages 폴더가 위치한 경로입니다.</param>
    /// <param name="targetLanguage">점검할 대상 언어입니다.</param>
    public static void CheckTranslation(string pathContainsLanguages, LanguageInfo targetLanguage)
    {
        Log.Msg("", true);
        Log.Msg("번역 기본 검증 시작.", true);
        Log.Msg($"\t{pathContainsLanguages}", true);
        
        // 1. Languages 폴더 확인
        var languagesPath = Path.Combine(pathContainsLanguages, "Languages");
        if (!Directory.Exists(languagesPath))
        {
            Log.Err("Languages 폴더를 찾을 수 없습니다.", true);
            Log.Err("번역 기본 검증 종료.", true);
            return;
        }
        Log.Good("Languages 폴더 확인.", true);

        // 2. 로케일 폴더 검색
        var foundLocaleFolder = CheckLocaleFolder(languagesPath, targetLanguage);
        if (string.IsNullOrEmpty(foundLocaleFolder))
        {
            Log.Err("번역 기본 검증 종료.", true);
        }

        // 3. XML 파일 텍스트 정밀 검증
        var xmlFiles = Directory.GetFiles(foundLocaleFolder, "*.xml", SearchOption.AllDirectories);
        if (xmlFiles.Length > 0)
        {
            var totalWarnings = CheckXmlFiles(xmlFiles);
            if (totalWarnings == 0)
            {
                Log.Good("XML 파일 검사 완료.", true);
                Log.Msg("\t잠재적 문제점을 발견하지 못했습니다.", true);
            }
            else if (totalWarnings > 0)
            {
                Log.Wrn("XML 파일 검사 완료.", true);
                Log.Msg($"\t{totalWarnings}개의 잠재적 문제점을 발견했습니다.", true);
            }
        }
        
        // 4. LoadFolders.Build.yaml 기초 정보 확인
        CheckBuildFile(pathContainsLanguages);

        Log.Good("번역 기본 검증 완료.", true);
        Log.Msg("\t기본 검증에서 확인하는 내용은 매우 기본적인 항목입니다.\n\t기본 검증에서 문제점을 발견하거나 하지 않았다고 해서 번역상 문제점이 반드시 있거나 없는 것은 아닙니다.\n\t기본 검증 내역은 사용자의 판단을 위한 참고 자료로 활용하기 바랍니다.", true);
    }
    
    private static void CheckBuildFile(string rootPath)
    {
        var yamlPath = Path.Combine(rootPath, "LoadFolders.Build.yaml");
        if (File.Exists(yamlPath))
        {
            Log.Good("빌드 파일 확인.", true);
            
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

            Log.Msg($"\tModName: {modName}", true);
            Log.Msg($"\tPackageID: {packageId}", true);
        }
        else
        {
            Log.Wrn("최상위 경로에 LoadFolders.Build.yaml 파일이 없습니다.\nRMK에 새로운 번역을 추가하고자 하는 경우 빌드 파일을 작성해주시기 바랍니다.", true);
        }
    }
    
    /// <summary>
    /// languagesDir에서 searchLanguage에 해당하는 언어 폴더가 존재하는지 확인합니다.
    /// </summary>
    /// <param name="languagesDir">검색할 Languages 폴더 경로입니다.</param>
    /// <param name="searchLanguage">존재 여부를 확인할 언어입니다.</param>
    /// <returns>확인에 성공하면 해당 경로를 반환하고, 실패하면 빈 문자열을 반환합니다.</returns>
    private static string CheckLocaleFolder(string languagesDir, LanguageInfo searchLanguage)
    {
        if (!Directory.Exists(languagesDir))
        {
            Log.Err($"경로가 존재하지 않습니다.\n\t{languagesDir}", true);
            return string.Empty;
        }
        var tentativeLocalePaths = Directory.GetDirectories(languagesDir);

        if (tentativeLocalePaths.Length == 0)
        {
            Log.Err($"Languages에 언어 폴더가 존재하지 않습니다.", true);
            return string.Empty;
        }

        string? foundLocalePath = tentativeLocalePaths.FirstOrDefault(path =>
        {
            var dirName = Path.GetFileName(path);
            return LanguageInfo.TryParse(dirName, out LanguageInfo? foundLanguage) && 
                   foundLanguage!.Equals(searchLanguage);
        });

        if (string.IsNullOrEmpty(foundLocalePath))
            Log.Err($"Languages에 {searchLanguage.ExtendedName} 폴더가 존재하지 않습니다.", true);
        else
            Log.Good($"{searchLanguage.ExtendedName} 폴더 확인.", true);
        return foundLocalePath!;
    }

    /// <summary>여러 XML 파일을 검사하고, 발견한 잠재적 문제점 개수를 반환합니다.</summary>
    /// <param name="xmlFiles">검사할 XML 파일 경로입니다. 경로의 유효성은 검증하지 않으니 유의하세요.</param>
    private static int CheckXmlFiles(IEnumerable<string> xmlFiles)
    {
        int totalWarnings = 0;
        
        foreach (var file in xmlFiles)
            totalWarnings += CheckXmlFile(file);
        
        return totalWarnings;
    }
    
    /// <summary>단일 XML 파일을 검사하고, 발견한 잠재적 문제점 개수를 반환합니다.</summary>
    private static int CheckXmlFile(string filePath)
    {
        int fileWarnings = 0;
        
        var lines = File.ReadAllLines(filePath);
        if (lines.Length == 0) return fileWarnings;

        // 로그 출력 시 파일 식별을 위해 "부모폴더/파일명" 형태로 조합 (예: Keyed/Messages.xml)
        var parentDir = Path.GetFileName(Path.GetDirectoryName(filePath)) ?? "";
        var displayPath = Path.Combine(parentDir, Path.GetFileName(filePath)).Replace('\\', '/');
        
        // [체크 1] 선언부 유효성 검사
        var firstLine = lines[0].Trim('\uFEFF', ' ', '\t', '\n', '\r');
        if (!firstLine.StartsWith("<?xml version=\"1.0\" encoding=\"utf-8\"?>", StringComparison.OrdinalIgnoreCase))
        {
            Log.Err($"\t{displayPath} Line 1: XML 선언부가 누락되었거나 표준(<?xml version=\"1.0\" encoding=\"utf-8\"?>)과 다릅니다.", true);
            fileWarnings++;
        }

        // [체크 2] XML 파싱
        // HtmlDocument를 이용한 XML 파싱이 유연, 강력, 간단해보여 도입해보았습니다.
        var doc = new HtmlDocument();
        doc.OptionOutputOriginalCase = true;
        doc.LoadHtml(string.Join("\n", lines));

        foreach (var error in doc.ParseErrors)
        {
            Log.Err($"\t{displayPath} Line {error.Line}: {error.Reason}", true);
            fileWarnings++;
        }

        // [체크 3] 최상위 노드 유효성
        var root = doc.DocumentNode.ChildNodes
            .FirstOrDefault(n => n.NodeType == HtmlNodeType.Element && n.Name != "?xml");
        if (root == null)
        {
            Log.Err($"\t{displayPath}: 최상위 LanguageData 노드를 찾을 수 없습니다.", true);
            fileWarnings++;
            return fileWarnings;
        }

        if (!string.Equals(root.Name, "LanguageData", StringComparison.OrdinalIgnoreCase))
        {
            Log.Err($"\t{displayPath} Line {root.Line}: 최상위 노드가 \"LanguageData\"가 아닌 \"{root.Name}\"입니다.", true);
            fileWarnings++;
        }

        void InspectTextNode(HtmlNode targetNode)
        {
            var rawText = targetNode.InnerHtml; 
            if (string.IsNullOrEmpty(rawText)) return;

            // li 노드인 경우 부모 이름을 포함하여 로그 식별성 향상 (예: rulesStrings의 <li>)
            var displayName = targetNode.Name == "li" ? $"{targetNode.ParentNode.Name}의 <li>" : targetNode.Name;

            // 화이트스페이스 검사
            if (rawText.StartsWith(" ") || rawText.EndsWith(" ") ||
                rawText.StartsWith("\n") || rawText.EndsWith("\n") ||
                rawText.StartsWith("\t") || rawText.EndsWith("\t"))
            {
                Log.Wrn($"\t{displayPath} Line {targetNode.Line}: 텍스트 앞뒤에 공백이나 줄바꿈이 있습니다. ({displayName})", true);
                fileWarnings++;
            }

            // 앰퍼샌드 검사
            if (UnescapedAmpersandRegex.IsMatch(rawText))
            {
                Log.Wrn($"\t{displayPath} Line {targetNode.Line}: 텍스트에 예약어 처리되지 않은 '&'가 있습니다. '&amp;'로 치환해주세요. ({targetNode.Name})", true);
                fileWarnings++;
            }

            var decodedText = targetNode.InnerText; 
        
            // 조사 검사 1
            if (BadParticleFormatRegex.IsMatch(decodedText))
            {
                Log.Wrn($"\t{displayPath} Line {targetNode.Line}: 조사가 잘못 쓰였을 가능성이 있습니다. '(을)를'과 같은 형태인지 확인해주세요. ({targetNode.Name})", true);
                fileWarnings++;
            }
            
            // 조사 검사 2
            if (SpaceBeforeParticleRegex.IsMatch(decodedText))
            {
                Log.Wrn($"\t{displayPath} Line {targetNode.Line}: 조사가 잘못 쓰였을 가능성이 있습니다. 자동 조사는 앞글자에 붙여써야 합니다. ({targetNode.Name})", true);
                fileWarnings++;
            }
            
            // 조사 검사 3
            if (SingleParticleAfterVarRegex.IsMatch(decodedText))
            {
                Log.Wrn($"\t{displayPath} Line {targetNode.Line}: 조사가 잘못 쓰였을 가능성이 있습니다. '(을)를'과 같은 형태인지 확인해주세요. ({targetNode.Name})", true);
                fileWarnings++;
            }
        }
        
        foreach (var node in root.ChildNodes.Where(n => n.NodeType == HtmlNodeType.Element))
        {
            // 1. 자식으로 li 태그가 있는지 확인하여 컨테이너 여부 판단
            var liNodes = node.ChildNodes
                .Where(c => c.NodeType == HtmlNodeType.Element && string.Equals(c.Name, "li", StringComparison.OrdinalIgnoreCase))
                .ToList();

            // 2. 텍스트 내부에 짝이 맞는 < > 가 잘못 사용된 경우 탐지 (li 이외의 뜬금없는 태그가 생성된 경우)
            var unexpectedTags = node.ChildNodes
                .Where(c => c.NodeType == HtmlNodeType.Element && !string.Equals(c.Name, "li", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (unexpectedTags.Any())
            {
                var tagNames = string.Join(", ", unexpectedTags.Select(t => $"<{t.Name}>"));
                Log.Err($"\t{displayPath} Line {node.Line}: 텍스트 내부에 예약어 처리되지 않은 괄호 짝(< >)이 포함되어 있습니다. 각각 '&lt;', '&gt;'로 치환해주세요. ({node.Name})", true);
                fileWarnings++;
            }

            // 3. 분기 처리: 컨테이너 노드 vs 단일 텍스트 노드
            if (liNodes.Any())
            {
                // 컨테이너 노드: 컨테이너 껍데기의 공백/줄바꿈은 정상으로 취급하고, 자식 <li> 들의 내용물만 정밀 검사
                foreach (var li in liNodes) InspectTextNode(li);
                
            }
            else
            {
                // 단일 텍스트 노드: 기존처럼 자기 자신 검사
                InspectTextNode(node);
            }
        }

        return fileWarnings;
    }
}