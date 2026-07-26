using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using RimworldExtractorInternal.Core;
using RimworldExtractorInternal.DataTypes;
using RimworldExtractorInternal;

namespace RimworldExtractorGUI.Services;

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
    Task<ExtractionResultSummary> ExtractAndSaveAsync(
        ModMetadata targetMod,
        List<ExtractableFolder> selectedFolders,
        List<ModMetadata> referenceMods);

    Task ConvertXmlToXlsxAsync(string[] rootFolderPaths);

    Task ConvertXlsxToXmlAsync(string xlsxFilePath);

    Task SaveRefModsListAsync(string filePath, IEnumerable<string> modIdentifiers);

    Task<List<string>> LoadRefModsListAsync(string filePath);
}

public class ExtractionService : IExtractionService
{
    public async Task<ExtractionResultSummary> ExtractAndSaveAsync(
        ModMetadata targetMod,
        List<ExtractableFolder> selectedFolders,
        List<ModMetadata> referenceMods)
    {
        var extraction = await Task.Run(() =>
        {
            // 1. 참조 모드들로부터 Defs 경로와 PrePatches 수집
            var refDefs = new List<string>();
            var prePatches = new List<ExtractableFolder>();

            if (referenceMods != null)
            {
                foreach (var referenceMod in referenceMods)
                {
                    refDefs.AddRange(
                        from extractableFolder in ModLister.GetExtractableFolders(referenceMod)
                        where extractableFolder.IsAutoSelectable() && Path.GetFileName(extractableFolder.FolderName) == "Defs"
                        select Path.Combine(referenceMod.RootDir, extractableFolder.FolderName)
                    );

                    prePatches.AddRange(
                        ModLister.GetExtractableFolders(referenceMod)
                            .Where(x => x.IsAutoSelectable() && Path.GetFileName(x.FolderName) == "Patches")
                    );
                }
            }

            // 2. DefTree 파이프라인 시뮬레이션 실행 (Roslyn PostProcessors 및 랭귀지 오버라이드 포함)
            var simResult = DefTreeSimulator.Execute(
                targetMod,
                selectedFolders,
                prePatches,
                refDefs,
                targetMod.IsOfficialContent,
                referenceMods);

            // 3. 시뮬레이션 결과(SimulationResult)를 인수로 수령하여 번역 항목 추출
            return Extractor.ExtractTranslationData(simResult);
        });

        var outPath = targetMod.Identifier.StripInvaildChars();

        await Task.Run(() =>
        {
            switch (ConfigManager.Current.Method)
            {
                // Prefabs.ExtractionMethod에서 ExtractionMethod로 직접 참조하도록 수정
                case ExtractionMethod.Excel:
                    IO.ToExcel(extraction, Path.Combine(outPath, outPath));
                    break;
                case ExtractionMethod.Languages:
                    IO.ToLanguageXml(extraction, false, false, outPath, outPath);
                    break;
                case ExtractionMethod.LanguagesWithComments:
                    IO.ToLanguageXml(extraction, false, true, outPath, outPath);
                    break;
            }

            string buildYamlText = RimworldExtractorInternal.Core.Utils.WriteBuildYamlText(targetMod);
            File.WriteAllText(Path.Combine(outPath, "LoadFolders.Build.yaml"), buildYamlText);
        });

        var (cntDefs, cntKeyed, cntStrings, cntPatches) = RimworldExtractorInternal.Core.Utils.Count(extraction);

        return new ExtractionResultSummary(
            TotalCount: extraction.Count,
            DefsCount: cntDefs,
            KeyedCount: cntKeyed,
            StringsCount: cntStrings,
            PatchesCount: cntPatches,
            OutputPath: outPath
        );
    }

    public async Task ConvertXmlToXlsxAsync(string[] rootFolderPaths)
    {
        await Task.Run(() =>
        {
            for (var i = 0; i < rootFolderPaths.Length; i++)
            {
                var root = rootFolderPaths[i];
                var translations = IO.FromLanguageXml(root);
                IO.ToExcel(translations, Path.Combine(root, Path.GetFileNameWithoutExtension(root)));
                Log.Msg($"{i + 1}/{rootFolderPaths.Length}:: 작업 완료: {root}");
            }
        });
    }

    public async Task ConvertXlsxToXmlAsync(string xlsxFilePath)
    {
        await Task.Run(() =>
        {
            var translations = IO.FromExcel(xlsxFilePath);
            // Prefabs.CommentOriginal을 ConfigManager.Current.CommentOriginal로 교체
            IO.ToLanguageXml(translations, true, ConfigManager.Current.CommentOriginal, Path.GetFileName(xlsxFilePath), Path.GetDirectoryName(xlsxFilePath) ?? "");
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