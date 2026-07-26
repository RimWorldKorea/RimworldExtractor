using ClosedXML.Excel;
using System.Xml.Linq;

namespace RimExtractorCore
{
    public static class ExtractorConstants
    {
        // 번역 엔트리 확장을 위한 키
        public const string ExtensionKeyExtraCommentTranslated = "ExtraCommentTranslated";
        
        // 파일 덮어쓰기 정책 처리를 위한 콜백 이벤트
        public static Action<XLWorkbook, string>? StopCallbackXlsx;
        public static Action<XDocument, string>? StopCallbackXml;
        public static Action<IEnumerable<string>, string>? StopCallbackTxt;
    }
}