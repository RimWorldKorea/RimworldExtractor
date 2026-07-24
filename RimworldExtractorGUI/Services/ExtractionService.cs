using RimworldExtractorInternal;
using RimworldExtractorInternal.DataTypes;

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
            Extractor.ExtractTranslationData(targetMod, selectedFolders, referenceMods));

        var outPath = targetMod.Identifier.StripInvaildChars();

        await Task.Run(() =>
        {
            switch (Prefabs.Method)
            {
                case Prefabs.ExtractionMethod.Excel:
                    IO.ToExcel(extraction, Path.Combine(outPath, outPath));
                    break;
                case Prefabs.ExtractionMethod.Languages:
                    IO.ToLanguageXml(extraction, false, false, outPath, outPath);
                    break;
                case Prefabs.ExtractionMethod.LanguagesWithComments:
                    IO.ToLanguageXml(extraction, false, true, outPath, outPath);
                    break;
            }

            string buildYamlText = RimworldExtractorInternal.Utils.WriteBuildYamlText(targetMod);
            File.WriteAllText(Path.Combine(outPath, "LoadFolders.Build.yaml"), buildYamlText);
        });

        var (cntDefs, cntKeyed, cntStrings, cntPatches) = RimworldExtractorInternal.Utils.Count(extraction);

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
            IO.ToLanguageXml(translations, true, Prefabs.CommentOriginal, Path.GetFileName(xlsxFilePath), Path.GetDirectoryName(xlsxFilePath) ?? "");
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