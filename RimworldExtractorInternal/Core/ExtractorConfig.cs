using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;

namespace RimworldExtractorInternal.Core
{
    public class ExtractorConfig
    {
        // 1. 기본 경로 및 버전 (기존 Prefabs의 기본값 복원)
        public bool EnableTkey { get; set; } = false;
        public string PathRimworld { get; set; } = "C:\\Program Files (x86)\\Steam\\steamapps\\common\\RimWorld";
        public string PathWorkshop { get; set; } = "C:\\Program Files (x86)\\Steam\\steamapps\\workshop\\content\\294100";
        public string PathBaseRefList { get; set; } = "";
        public string CurrentVersion { get; set; } = "1.6";
        public string PatternVersion { get; set; } = @"^[1]\.\d+";
        public string PatternVersionWithV { get; set; } = @"^v[1]\.\d+";

        // 2. 언어 및 출력 설정
        public string OriginalLanguage { get; set; } = "English";
        public string SecondaryLanguage { get; set; } = string.Empty;
        public string TranslationLanguage { get; set; } = "Korean (한국어)";
        public bool CommentOriginal { get; set; } = false;

        // 3. 추출 규칙 (JSON 직렬화 대상)
        public HashSet<string> ExtractableTags { get; set; } = new();
        public HashSet<string> FullListTranslationTags { get; set; } = new();
        public Dictionary<string, string> NodeReplacement { get; set; } = new();
        public List<string> TranslationHandles { get; set; } = new();

        [JsonConverter(typeof(JsonStringEnumConverter))]
        public DuplicatesPolicy Policy { get; set; } = DuplicatesPolicy.Overwrite;

        [JsonConverter(typeof(JsonStringEnumConverter))]
        public ExtractionMethod Method { get; set; } = ExtractionMethod.Languages;

        // 4. 내부 로직 (직렬화 무시)
        [JsonIgnore]
        private Dictionary<string, ExtractionRule>? _extractionRulesCache;

        public bool CanExtract(string tagName, string defName)
        {
            if (_extractionRulesCache == null)
            {
                _extractionRulesCache = new Dictionary<string, ExtractionRule>();
                foreach (var item in ExtractableTags)
                {
                    var parsedRule = new ExtractionRule(item);
                    _extractionRulesCache[parsedRule.Tag] = parsedRule;
                }
            }
            return _extractionRulesCache.TryGetValue(tagName, out var cachedRule) && cachedRule.CanExtract(defName);
        }

        public void InvalidateCache()
        {
            _extractionRulesCache = null;
        }

        public IEnumerable<string> GetLanguagePriorityList()
        {
            var list = new List<string>();
            if (!string.IsNullOrWhiteSpace(OriginalLanguage)) list.Add(OriginalLanguage);
            if (!string.IsNullOrWhiteSpace(SecondaryLanguage)) list.Add(SecondaryLanguage);
            if (!list.Contains("English")) list.Add("English");
            return list.Distinct();
        }
    }

    // 기존 Prefabs 하위에 있던 Enum 분리
    public enum DuplicatesPolicy { Stop = 0, Overwrite, KeepOriginal }
    public enum ExtractionMethod { Excel = 0, Languages, LanguagesWithComments }

    // 기존 Prefabs.ExtractionRule 완벽 복원
    public class ExtractionRule
    {
        public string Tag;
        public HashSet<string> Whitelist = new();
        public HashSet<string> Blacklist = new();

        public ExtractionRule(string raw)
        {
            var plusIndex = raw.IndexOf('+');
            var minusIndex = raw.IndexOf('-');
            if (plusIndex == -1 && minusIndex == -1)
            {
                Tag = raw;
                return;
            }
            var firstSepIndex = (plusIndex != -1 && minusIndex != -1)
                ? Math.Min(plusIndex, minusIndex)
                : Math.Max(plusIndex, minusIndex);
            
            Tag = raw.Substring(0, firstSepIndex);
            var remain = raw.Substring(firstSepIndex);
            int i = 0;
            
            while (i < remain.Length)
            {
                char mode = remain[i];
                int nextPlus = remain.IndexOf('+', i + 1);
                int nextMinus = remain.IndexOf('-', i + 1);
                int nextSep = (nextPlus == -1 && nextMinus == -1) ? remain.Length :
                              (nextPlus == -1) ? nextMinus :
                              (nextMinus == -1) ? nextPlus :
                              Math.Min(nextPlus, nextMinus);
                
                var content = remain.Substring(i + 1, nextSep - (i + 1));
                var items = content.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                var targetSet = (mode == '+') ? Whitelist : Blacklist;
                foreach (var item in items) targetSet.Add(item.Trim());
                i = nextSep;
            }
        }

        public bool CanExtract(string defName)
        {
            if (Whitelist.Count > 0)
            {
                if (!Whitelist.Contains(defName)) return false;
            }
            if (Blacklist.Contains(defName)) return false;
            return true;
        }

        public override string ToString()
        {
            var sb = new StringBuilder(Tag);
            if (Whitelist.Count > 0)
            {
                sb.Append("+");
                sb.Append(string.Join(",", Whitelist));
            }
            if (Blacklist.Count > 0)
            {
                sb.Append("-");
                sb.Append(string.Join(",", Blacklist));
            }
            return sb.ToString();
        }
    }
}