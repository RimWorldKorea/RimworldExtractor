using System.Xml.Linq;

namespace RimExtractorCore
{
    public static class Constants
    {
        // 번역 관련 어트리뷰트 태그명 상수 정의
        public const string AttrNoTranslate = "NoTranslate";
        public const string AttrMayNotTranslate = "MayNotTranslate";
        public const string AttrMayTranslate = "MayTranslate";
        public const string AttrMustTranslate = "MustTranslate";
        
        // [NEW] C# 어트리뷰트 이름 -> XML 태그(어트리뷰트) 이름 매핑 테이블
        public static readonly Dictionary<string, string> TranslationAttributes = new(StringComparer.OrdinalIgnoreCase)
        {
            { "NoTranslateAttribute", AttrNoTranslate },
            { "TranslationMayNotNecessaryAttribute", AttrMayNotTranslate },
            { "MustTranslateAttribute", AttrMustTranslate }
        };
        
        // 번역 엔트리 확장을 위한 키
        public const string ExtensionKeyExtraCommentTranslated = "ExtraCommentTranslated";
        
        // 파일 덮어쓰기 정책 처리를 위한 콜백 이벤트
        public static Action<XDocument, string>? StopCallbackXml;
        public static Action<IEnumerable<string>, string>? StopCallbackTxt;
    }
}