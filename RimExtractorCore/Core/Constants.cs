using System.Xml.Linq;

namespace RimExtractorCore;

public static class Constants
{
    // XML 번역 어트리뷰트 태그 정의 (추출기 자체적으로 사용하는 태그입니다.)
    public const string AttrNoTranslate = "NoTranslate";
    public const string AttrMayNotTranslate = "MayNotTranslate"; // 추출기에서 정의한 고유 태그
    public const string AttrMayTranslate = "MayTranslate";
    public const string AttrMustTranslate = "MustTranslate";
    public const string AttrTranslationCanChangeCount = "FullListTranslate";
    
    /// <summary>
    /// 림월드 어트리뷰트 이름을 자체 XML 태그로 연결하기 위한 테이블입니다.
    /// DefTreeSimulator에서 이 테이블을 참조합니다.
    /// </summary>
    public static readonly Dictionary<string, string> TranslationAttributes = new(StringComparer.OrdinalIgnoreCase)
    {
        { "NoTranslateAttribute", AttrNoTranslate },
        { "MayTranslateAttribute", AttrMayTranslate },
        //{ "TranslationMayNotNecessaryAttribute", AttrMayNotTranslate }, <- 자체 추가 태그라 어셈블리에 애초에 없음
        { "MustTranslateAttribute", AttrMustTranslate },
        { "TranslationCanChangeCountAttribute", AttrTranslationCanChangeCount },
    };

    // 번역 엔트리 확장을 위한 키
    public const string ExtensionKeyExtraCommentTranslated = "ExtraCommentTranslated";

    // 파일 덮어쓰기 정책 처리를 위한 콜백 이벤트
    public static Action<XDocument, string>? StopCallbackXml;
    public static Action<IEnumerable<string>, string>? StopCallbackTxt;
}