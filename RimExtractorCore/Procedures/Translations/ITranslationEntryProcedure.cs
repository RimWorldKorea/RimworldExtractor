using RimExtractorCore.DataTypes;

namespace RimExtractorCore.Procedures;

/// <summary>
/// Extractor가 추출한 TranslationEntry 데이터를 후가공하는 프로시저들의 공통 인터페이스입니다.
/// </summary>
public interface ITranslationEntryProcedure
{
    string Name { get; }

    IEnumerable<TranslationEntry> Process(IEnumerable<TranslationEntry> entries);
}