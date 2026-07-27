using System.Collections.Generic;
using System.Linq;
using RimExtractorCore.DataTypes;
using RimExtractorCore;

namespace RimExtractorCore.Procedures;

public class NodeReplacementProcedure : ITranslationEntryProcedure
{
    public string Name => "NodeReplacementProcedure";
    
    //TODO 리팩토링용 임시
    private readonly Dictionary<string, string> NodeReplacements = new()
    {
        { "CombatExtended.AmmoDef+*", "ThingDef+*" },
        { "VFECore.ExpandableProjectileDef+*", "ThingDef+*" },
        { "AbilityUser.ProjectileDef_AbilityLaser+*", "ThingDef+*" },
        { "AbilityUser.ProjectileDef_Ability+*", "ThingDef+*" },
        { "NewRatkin.CustomThingDef+*", "ThingDef+*" },
        { "AlienRace.AlienBackstoryDef+*", "BackstoryDef+*" },
        { "RatkinGeneExpanded.FactionDefExtended+*", "FactionDef+*" },
        { "RatkinGeneExpanded.ThingDefExtended+*", "ThingDef+*" },
        { "AlienRace.ThingDef_AlienRace+*", "ThingDef+*" },
        { "Rimlaser.Building_LaserGunDef+*", "ThingDef+*" },
        { "Rimlaser.LaserBeamDef+*", "ThingDef+*" },
        { "Rimlaser.LaserGunDef+*", "ThingDef+*" },
        { "Rimlaser.SpinningLaserGunDef+*", "ThingDef+*" },
        { "JecsTools.BackstoryDef+baseDesc", "JescTools.BackstoryDef+description" },
        { "AnestheticGunMod2.AnestheticBulletDef+*", "ThingDef+*" },
        { "BackstoryDef+baseDesc", "BackstoryDef+description" },
        { "DubsBadHygiene.WashingJobDef+*", "JobDef+*" },
        { "DubsBadHygiene.Needy+*", "NeedDef+*" },
        { "VarietyMatters.FoodVariety_NeedDef+*", "NeedDef+*" },
        { "Kiiro.StorytellerDef_Custom+*", "StorytellerDef+*" },
        { "Vehicles.SkinDef+*", "Vehicles.PatternDef+*" },
        { "Vehicles.AntiAircraftDef+*", "WorldObjectDef+*" },
        { "Vehicles.AirdropDef+*", "ThingDef+*" },
        { "Meow.FactionDefExtended+*", "FactionDef+*" }
    };

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
        
        foreach (var (key, value) in this.NodeReplacements)
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