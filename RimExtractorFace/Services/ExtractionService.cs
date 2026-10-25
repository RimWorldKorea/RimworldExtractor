using RimExtractorCore;
using RimExtractorCore.Extractor;
using RimExtractorCore.DataTypes;
using RimExtractorCore.DefTreeSimulator;

namespace RimExtractorFace.Services;

public record ExtractionResultSummary(
    int TotalCount,
    int DefsCount,
    int KeyedCount,
    int StringsCount,
    int PatchesCount,
    string OutputPath
);

public interface IExtractionService
{
    Task<ExtractionResultSummary> ExtractAndSaveAsync(Mod targetMod, List<Mod> referenceMods);
    Task ConvertXmlToOdsAsync(string[] rootFolderPaths);
    Task ConvertSheetToXmlAsync(string xlsxFilePath);
    Task SaveRefModsListAsync(string filePath, IEnumerable<string> modIdentifiers);
    Task<List<string>> LoadRefModsListAsync(string filePath);
}

public class ExtractionService : IExtractionService
{
    public async Task<ExtractionResultSummary> ExtractAndSaveAsync(Mod targetMod, List<Mod> referenceMods)
    {
        var extraction = await Task.Run(() =>
        {
            if (!Version.TryParse(SettingManager.Current.CurrentVersion, out var targetVersion))
                throw new ArgumentException($"현재 버전 설정이 올바르지 않습니다: {SettingManager.Current.CurrentVersion}");

            // 새로운 DefTreeSimulator 시그니처에 맞게 호출
            var simResult = DefTreeSimulator.Execute(targetMod, targetVersion, referenceMods);
            return ExtractorEngine.ExtractTranslationData(simResult);
        });

        // ModMetadata.Identifier 대체 -> Name과 WorkshopID 조합 혹은 PackageID 사용
        var modIdentifier = targetMod.Type == ModType.Official ? targetMod.Name : $"{targetMod.Name} - {targetMod.WorkshopID}";
        var outPath = modIdentifier?.StripInvaildChars() ?? "UnknownMod";

        await Task.Run(() =>
        {
            switch (SettingManager.Current.Method)
            {
                case ExportFileFormmat.Spreadsheet:
                    FileInterface.ToOds(extraction.Entries.ToList(), Path.Combine(outPath, outPath));
                    break;
                case ExportFileFormmat.LanguageData:
                    FileInterface.ToLanguageXml(extraction.Entries.ToList(), false, false, outPath, outPath);
                    break;
                case ExportFileFormmat.LanguageDataWithComments:
                    FileInterface.ToLanguageXml(extraction.Entries.ToList(), false, true, outPath, outPath);
                    break;
            }

            string buildYamlText = RimExtractorCore.Utils.WriteBuildYamlText(targetMod);
            File.WriteAllText(Path.Combine(outPath, "LoadFolders.Build.yaml"), buildYamlText);
        });

        var (cntDefs, cntKeyed, cntStrings, cntPatches) = RimExtractorCore.Utils.Count(extraction.Entries);

        return new ExtractionResultSummary(
            extraction.Entries.Count, cntDefs, cntKeyed, cntStrings, cntPatches, outPath
        );
    }

    public async Task ConvertXmlToOdsAsync(string[] rootFolderPaths)
    {
        await Task.Run(() =>
        {
            for (var i = 0; i < rootFolderPaths.Length; i++)
            {
                var root = rootFolderPaths[i];
                var translations = FileInterface.FromLanguageXml(root);
                FileInterface.ToOds(translations, Path.Combine(root, Path.GetFileNameWithoutExtension(root)));
                Log.Msg($"{i + 1}/{rootFolderPaths.Length}:: 작업 완료: {root}");
            }
        });
    }

    public async Task ConvertSheetToXmlAsync(string xlsxFilePath)
    {
        await Task.Run(() =>
        {
            var translations = FileInterface.FromSheetFile(xlsxFilePath);
            FileInterface.ToLanguageXml(
                translations, 
                true, 
                SettingManager.Current.CommentOriginal, 
                Path.GetFileName(xlsxFilePath), 
                Path.GetDirectoryName(xlsxFilePath) ?? "");
        });
    }

    public async Task SaveRefModsListAsync(string filePath, IEnumerable<string> modIdentifiers)
    {
        await File.WriteAllLinesAsync(filePath, modIdentifiers);
    }

    public async Task<List<string>> LoadRefModsListAsync(string filePath)
    {
        if (!File.Exists(filePath)) return new List<string>();
        var lines = await File.ReadAllLinesAsync(filePath);
        return lines.ToList();
    }
}