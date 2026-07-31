using System.Collections.Generic;
using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.Procedures

{
    /// <summary>
    /// 단일 DefTree를 순회하여 번역 데이터를 추출합니다.
    /// </summary>
    public interface IExtractionProcedure
    {
        string Name { get; }
        IEnumerable<TranslationEntry> Extract(DefSnapshot snapshot, ModMetadata targetMod);
    }
}