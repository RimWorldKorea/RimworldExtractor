using System.Collections.Generic;
using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorCore.Procedures

{
    /// <summary>
    /// 단일 DefSnapshot(우주 1개)의 XML 트리를 순회하여 번역 데이터를 추출합니다.
    /// 차분(Diff) 계산은 코어 엔진이 알아서 처리하므로, 여기서는 단순히 다 뽑아내기만 하면 됩니다.
    /// </summary>
    public interface IExtractionProcedure
    {
        string Name { get; }
        IEnumerable<TranslationEntry> Extract(DefSnapshot snapshot, ModMetadata targetMod);
    }
}