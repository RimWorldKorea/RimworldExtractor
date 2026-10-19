using System.Collections.Immutable;
using System.Xml.Linq;

namespace RimExtractorCore;

public static class Constants
{
    /// <summary>림월드 GenTypes에서 네임스페이스 명시 없이도 접근을 허용하는 바닐라 VIP 명단입니다.
    /// 다시 말해, 이 네임스페이스 외부에서 정의된 커스텀 Def는 네임스페이스를 붙여서 참조해야 합니다.</summary>
    public static readonly HashSet<string> IgnoredNamespaceNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "RimWorld", "Verse", "LudeonTK", "Verse.AI", "Verse.AI.Group", "Verse.Sound", "Verse.Grammar",
        "RimWorld.Planet", "RimWorld.BaseGen", "RimWorld.QuestGen", "RimWorld.SketchGen", "System"
    };
    
    // XML 번역 어트리뷰트 태그 정의 (추출기 자체적으로 사용하는 태그입니다.)
    public const string AttrNoTranslate = "NoTranslate";
    public const string AttrMayNotTranslate = "MayNotTranslate"; // 추출기에서 정의한 고유 태그
    public const string AttrMayTranslate = "MayTranslate";
    public const string AttrMustTranslate = "MustTranslate";
    public const string AttrTranslationCanChangeCount = "FullListTranslate";
    
    /// <summary>림월드 어트리뷰트 이름을 자체 XML 태그로 연결하기 위한 테이블입니다.</summary>
    public static readonly Dictionary<string, string> TranslationAttributes = new(StringComparer.OrdinalIgnoreCase)
    {
        { "NoTranslateAttribute", AttrNoTranslate },
        { "MayTranslateAttribute", AttrMayTranslate },
        { "MustTranslateAttribute", AttrMustTranslate },
        { "TranslationCanChangeCountAttribute", AttrTranslationCanChangeCount },
    };
    
    /// <summary>번역 엔트리 확장을 위한 키</summary>
    /// <todo>이거 어디다 쓰는거지?</todo>
    public const string ExtensionKeyExtraCommentTranslated = "ExtraCommentTranslated";
    
    // 파일 덮어쓰기 정책 처리를 위한 콜백 이벤트
    //TODO 이거 어디 쓰는거지?
    public static Action<XDocument, string>? StopCallbackXml;
    public static Action<IEnumerable<string>, string>? StopCallbackTxt;
}