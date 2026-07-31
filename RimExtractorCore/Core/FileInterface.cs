using System.Xml;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore.Extractor;
using RimExtractorCore.Spreadsheet;

namespace RimExtractorCore
{
    public static class FileInterface
    {
        private static readonly string HeaderClassNode = "Class+Node [(Identifier (Key)]";
        private static readonly string HeaderClass = "Class [Not chosen]";
        private static readonly string HeaderNode = "Node [Not chosen]";
        private static readonly string HeaderRequiredMods = "Required Mods [Not chosen]";
        private static string HeaderOriginal => $"{SettingManager.Current.OriginalLanguage} [Source string]";
        private static string HeaderTranslated => $"{SettingManager.Current.TranslationLanguage} [Translation]";

        public static List<TranslationEntry> FromExcel(string inputPath)
        {
            return SpreadsheetReader.ReadTranslations(inputPath);
        }
        
        public static List<TranslationEntry> FromLanguageXml(string rootPath, bool isOfficialContent = false)
        {
            var translationsDir = Path.Combine(rootPath, "Languages", SettingManager.Current.TranslationLanguage);
            if (!Directory.Exists(translationsDir))
                translationsDir = Path.Combine(rootPath, "Languages", SettingManager.Current.TranslationLanguage.Split(' ').First());

            var defInjectedDir = Path.Combine(translationsDir, "DefInjected");
            var keyedDir = Path.Combine(translationsDir, "Keyed");
            var stringsDir = Path.Combine(translationsDir, "Strings");
            
            var translations = new List<TranslationEntry>();

            // 1. DefInjected 읽기 (I/O는 IO.cs가, 파싱은 LanguageXmlProcessor가 담당)
            foreach (var filePath in DescendantFiles(defInjectedDir).Where(x => x.ToLower().EndsWith(".xml")))
            {
                var className = Path.GetRelativePath(defInjectedDir, filePath).Split(Path.DirectorySeparatorChar).First();
                try
                {
                    var doc = ReadXml(filePath);
                    translations.AddRange(LanguageXmlProcessor.ParseDefInjected(doc, className));
                }
                catch (Exception e)
                {
                    Log.Err($"{filePath} 읽기 실패: {e.Message}");
                    throw;
                }
            }

            // 2. Keyed 읽기 (ExtractorEngine 활용)
            var keyed = new ExtractableFolder(ModMetadata.Emptry, keyedDir, null);
            translations.AddRange(ExtractorEngine.ExtractKeyed(keyed, isOfficialContent)
                .Select(x => x with { Translated = x.Original, Original = "" }));

            // 3. Strings 읽기 (ExtractorEngine 활용)
            var strings = new ExtractableFolder(ModMetadata.Emptry, stringsDir, null);
            translations.AddRange(ExtractorEngine.ExtractStrings(strings)
                .Select(x => x with { Translated = x.Original, Original = "" }));

            return translations;
        }
        
        public static void ToLanguageXml(List<TranslationEntry> translations, bool skipNoTranslation, bool commentOriginal, string ModName, string rootDirPath)
        {
            // 1. LanguageXmlProcessor를 통해 파일 생성에 필요한 경로와 데이터를 모두 메모리 상에서 완성하여 받아옵니다.
            var (xmlFiles, txtFiles) = LanguageXmlProcessor.GenerateLanguageFiles(
                translations, skipNoTranslation, commentOriginal, ModName, rootDirPath);

            // 2. 반환받은 Dictionary를 순회하며 순수하게 '폴더 생성'과 '파일 저장'만 수행합니다. (XML)
            foreach (var (path, doc) in xmlFiles)
            {
                var dir = Path.GetDirectoryName(path);
                if (dir != null && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
        
                doc.SaveSafely(path);
            }

            // 3. 텍스트 파일 저장 (Strings)
            foreach (var (path, lines) in txtFiles)
            {
                var dir = Path.GetDirectoryName(path);
                if (dir != null && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
        
                lines.SaveSafely(path);
            }
        }

        /// <summary>
        /// 추출된 번역 데이터를 ODS 파일로 저장합니다.
        /// </summary>
        public static void ToOds(List<TranslationEntry> translations, string outPath)
        {
            var grid = new Grid { Name = "Translations" };

            // 1. 헤더 추가
            grid.AppendRow(new[]
            {
                HeaderClassNode, HeaderClass, HeaderNode, HeaderRequiredMods, HeaderOriginal, HeaderTranslated
            });

            // 2. 데이터 추가
            grid.AppendRows(translations.Select(t => t.ToGridRow()));

            // 3. ODS 저장 (System.IO.Compression 기반 자체 로직)
            var fullPath = outPath.EndsWith(".ods", StringComparison.OrdinalIgnoreCase) 
                ? outPath : outPath + ".ods";
                
            OdsWriter.SaveGrid(fullPath, grid);
        }
        
        private static void SaveSafely(this XDocument doc, string path)
        {
            if (!File.Exists(path))
            {
                doc.Save(path);
                return;
            }

            switch (SettingManager.Current.Policy)
            {
                case DuplicatesPolicy.Stop:
                    var stopCallback = Constants.StopCallbackXml;
                    if (stopCallback != null)
                        stopCallback(doc, path);
                    else
                        throw new ArgumentNullException(nameof(stopCallback));
                    return;
                case DuplicatesPolicy.Overwrite:
                    doc.Save(path);
                    return;
                case DuplicatesPolicy.KeepOriginal:
                    return;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static void SaveSafely(this IEnumerable<string> lines, string path)
        {
            if (!File.Exists(path))
            {
                File.WriteAllLines(path, lines);
                return;
            }

            switch (SettingManager.Current.Policy)
            {
                case DuplicatesPolicy.Stop:
                    var stopCallback = Constants.StopCallbackTxt;
                    if (stopCallback != null)
                        stopCallback(lines, path);
                    else
                        throw new ArgumentNullException(nameof(stopCallback));
                    return;
                case DuplicatesPolicy.Overwrite:
                    File.WriteAllLines(path, lines);
                    return;
                case DuplicatesPolicy.KeepOriginal:
                    return;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static XDocument ReadXml(string filePath)
        {
            var readerSettings = new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true,
                CheckCharacters = false
            };
    
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var xmlReader = XmlReader.Create(stream, readerSettings);
            return XDocument.Load(xmlReader);
        }

        public static IEnumerable<string> DescendantFiles(string root)
        {
            if (!Directory.Exists(root))
                yield break;

            var q = new Queue<string>();
            q.Enqueue(root);

            while (q.Count > 0)
            {
                var curPath = q.Dequeue();
                foreach (var subDir in Directory.GetDirectories(curPath).OrderBy(x => x))
                {
                    q.Enqueue(subDir);
                }

                foreach (var file in Directory.EnumerateFiles(curPath).OrderBy(x => x))
                {
                    yield return file;
                }
            }
        }
        
        
    }
}