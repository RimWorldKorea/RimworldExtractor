using System.Collections.Generic;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorInternal;

/// <summary>
/// XML 트리에서 추출이 완료된 번역 항목(TranslationEntry) 리스트를 가공하는 후처리 절차를 정의합니다.
/// </summary>
public interface ITranslationProcedure
{
    string Name { get; }

    IEnumerable<TranslationEntry> Process(IEnumerable<TranslationEntry> entries);
}