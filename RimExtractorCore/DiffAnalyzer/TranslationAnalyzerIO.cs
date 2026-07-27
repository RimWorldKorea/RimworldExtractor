using RimExtractorCore.DataTypes;

namespace RimExtractorCore.DiffAnalyzer;

public static class TranslationAnalyzerIO
{
    /// <summary>
    /// ODS 기반의 번역 파일 병합 및 저장 로직.
    /// 스타일링(셀 색상)이 배제되었으므로, 순수 데이터 갱신 후 FileInterface를 통해 덮어씁니다.
    /// </summary>
    public static void ModifySpreadsheet(TranslationAnalyzerEntry entry)
    {
        if (entry.NewTranslations == null)
        {
            throw new InvalidOperationException("병합할 새 번역 데이터가 없습니다.");
        }

        // 1. 메모리 상에서 번역 데이터 병합 (TranslationAnalyzerEntry의 내장 로직 활용)
        // 기존의 Translated 값을 신규 추출 데이터(NewTranslations)에 매핑
        entry.MergeTranslation();

        // 2. 저장 방식(SaveMethod)에 따른 경로 설정
        string outputPath = entry.FilePath;
        switch (entry.SaveMethod)
        {
            case TranslationAnalyzerEntry.SaveMethodEnum.Append:
            case TranslationAnalyzerEntry.SaveMethodEnum.Overwrite:
                outputPath = entry.FilePath;
                break;
            case TranslationAnalyzerEntry.SaveMethodEnum.RewriteNewFile:
            case TranslationAnalyzerEntry.SaveMethodEnum.New:
                // 필요시 새로운 파일명 규칙 적용 (예: _new.ods)
                var dir = Path.GetDirectoryName(entry.FilePath) ?? string.Empty;
                var fileName = Path.GetFileNameWithoutExtension(entry.FilePath);
                outputPath = Path.Combine(dir, $"{fileName}_Updated.ods");
                break;
        }

        // 3. 병합된 최종 데이터를 ODS로 바로 저장 (FileInterface 재활용)
        FileInterface.ToOds(entry.NewTranslations, outputPath);
        
        Log.Msg($"번역 파일 병합 완료: {Path.GetFileName(outputPath)}");
    }
}