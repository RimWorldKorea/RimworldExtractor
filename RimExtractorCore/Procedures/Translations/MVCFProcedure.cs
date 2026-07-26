using System.Collections.Generic;
using System.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore;


namespace RimExtractorCore.Procedures;

public class MVCFProcedure : ITranslationEntryProcedure
{
    public string Name => "MVCFProcedure";

    private const string verbPropsKeyword = "Comp_VerbProps.verbProps";
    private const string verbKeyword = ".verbs.";

    public IEnumerable<TranslationEntry> Process(IEnumerable<TranslationEntry> entries)
    {
        var translations = entries.ToList();
        var mvcfForms = new HashSet<MVCFForm>();

        foreach (var entry in translations.Where(x => x.Node.Contains(verbPropsKeyword)))
        {
            var tokens = entry.Node.Split('.');
            if (tokens.Last() == "label")
            {
                mvcfForms.Add(new MVCFForm(tokens[0], entry));
            }
        }

        foreach (var entry in translations.Where(x => x.Node.Contains(verbPropsKeyword)))
        {
            var tokens = entry.Node.Split('.');
            var lastToken = tokens.Last();
            if (lastToken == "visualLabel")
            {
                tokens[^1] = "label";
                var mvcfForm = mvcfForms.FirstOrDefault(x => x.VerbPropsLabel.Node == string.Join('.', tokens));
                if (mvcfForm != null) mvcfForm.VerbPropsVisualLabel = entry;
            }
            if (lastToken == "description")
            {
                tokens[^1] = "label";
                var mvcfForm = mvcfForms.FirstOrDefault(x => x.VerbPropsLabel.Node == string.Join('.', tokens));
                if (mvcfForm != null) mvcfForm.VerbPropsDescription = entry;
            }
        }

        foreach (var entry in translations.Where(x => x.Node.Contains(verbKeyword)))
        {
            var tokens = entry.Node.Split('.');
            if (tokens.Last() == "label")
            {
                var mvcfForm = mvcfForms.FirstOrDefault(x => x.DefName == tokens[0] && x.VerbPropsLabel.Original == entry.Original);
                if (mvcfForm != null) mvcfForm.VerbLabel = entry;
            }
        }

        foreach (var entry in translations)
        {
            var mvcfForm = mvcfForms.FirstOrDefault(x => x.VerbPropsLabel == entry);
            if (mvcfForm != null)
            {
                if (mvcfForm.VerbPropsVisualLabel == null)
                {
                    var tokens = mvcfForm.VerbPropsLabel.Node.Split('.');
                    tokens[^1] = "visualLabel";
                    yield return mvcfForm.VerbPropsLabel with { Node = string.Join('.', tokens) };
                }
                else yield return mvcfForm.VerbPropsVisualLabel;

                if (mvcfForm.VerbPropsDescription == null)
                {
                    var tokens = mvcfForm.VerbPropsLabel.Node.Split('.');
                    tokens[^1] = "description";
                    var verbPropsDescription = (new TranslationEntry(mvcfForm.VerbPropsLabel) with { Node = string.Join('.', tokens), Original = "" });
                    
                    // 💡 여기서 Prefabs 참조를 ExtractorConstants로 수정했습니다!
                    verbPropsDescription.AddExtension(ExtractorConstants.ExtensionKeyExtraCommentTranslated, "이 항목은 gizmo에 표시될 수 있습니다.");
                    
                    yield return verbPropsDescription;
                }
                else yield return mvcfForm.VerbPropsDescription;
            }

            if (mvcfForms.Any(x => x.VerbPropsLabel == entry || x.VerbLabel == entry || x.VerbPropsVisualLabel == entry))
                continue;

            yield return entry;
        }
    }

    private class MVCFForm
    {
        public readonly string DefName;
        public readonly TranslationEntry VerbPropsLabel;
        public TranslationEntry? VerbPropsVisualLabel = null;
        public TranslationEntry? VerbLabel = null;
        public TranslationEntry? VerbPropsDescription = null;

        public MVCFForm(string defName, TranslationEntry verbPropsLabel)
        {
            DefName = defName;
            VerbPropsLabel = verbPropsLabel;
        }
    }
}