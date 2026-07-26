using System.Collections.Generic;
using System.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore;

namespace RimExtractorCore.Procedures;

public class NodeReplacementProcedure : ITranslationEntryProcedure
{
    public string Name => "NodeReplacementProcedure";

    public IEnumerable<TranslationEntry> Process(IEnumerable<TranslationEntry> entries)
    {
        foreach (var entry in entries)
        {
            if (entry.ClassName is "Keyed" or "Strings")
                yield return entry;
            else
                yield return DoNodeReplacement(entry);
        }
    }

    private TranslationEntry DoNodeReplacement(TranslationEntry entry)
    {
        var isPatches = entry.ClassName.StartsWith("Patches.");
        var defType = isPatches ? entry.ClassName[("Patches.".Length + 1)..] : entry.ClassName;
        var defName = entry.Node.Split('.')[0];
        var nodeAfterDefName = entry.Node[(entry.Node.IndexOf('.') + 1)..];

        // 💡 Prefabs.NodeReplacement를 ConfigManager.Current.NodeReplacement로 교체했습니다.
        foreach (var (key, value) in ConfigManager.Current.NodeReplacement)
        {
            var tokenKey = key.Split('+');
            var tokenValue = value.Split("+");
            var targetDef = tokenKey[0];
            var targetNode = tokenKey[1] == "*" ? nodeAfterDefName : tokenKey[1];
            var changedDef = tokenValue[0];
            var changedNode = tokenValue[1] == "*" ? nodeAfterDefName : tokenValue[1];

            if (defType == targetDef && nodeAfterDefName == targetNode)
            {
                return entry with { ClassName = isPatches ? $"Patches.{changedDef}" : changedDef, Node = $"{defName}.{changedNode}" };
            }
        }
        return entry;
    }
}