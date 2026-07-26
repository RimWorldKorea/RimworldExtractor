using System.Text.Json;
using System.Text.RegularExpressions;

namespace RimExtractorCore
{
    public static class ConfigManager
    {
        public const string ConfigFileName = "config.json";
        
        public static ExtractorConfig Current { get; private set; } = new();

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions 
        { 
            WriteIndented = true 
        };

        public static void Load()
        {
            if (!File.Exists(ConfigFileName))
            {
                InitDefault();
                Save();
                return;
            }

            try
            {
                var jsonString = File.ReadAllText(ConfigFileName);
                Current = JsonSerializer.Deserialize<ExtractorConfig>(jsonString, JsonOptions) ?? new ExtractorConfig();
                Current.InvalidateCache();
            }
            catch (Exception ex)
            {
                Log.Err($"설정 파일 로드 실패. 기본값으로 덮어씁니다: {ex.Message}");
                InitDefault();
                Save();
            }
        }

        public static void Save()
        {
            try
            {
                Current.InvalidateCache();
                var jsonString = JsonSerializer.Serialize(Current, JsonOptions);
                File.WriteAllText(ConfigFileName, jsonString);
            }
            catch (Exception ex)
            {
                Log.Err($"설정 저장 실패: {ex.Message}");
            }
        }

        public static void InitDefault()
        {
            Current = new ExtractorConfig
            {
                // 기존 Prefabs.Init()에 있던 모든 기본 태그 복원
                ExtractableTags = new HashSet<string>(
                    "label/rulesStrings/description/baseDesc/title/titleShort/customLabel/symbol/jobString/reportString/labelNoun/slateRef/verb/gerund/adjective/member/tips/ideoName/thoughtStageDescriptions/jobReportString/theme/labelShortAdj/labelPlural/letterText/deathMessage/labelShort/letterLabel/helpText/text/baseInspectLine/labelFemale/descriptionShort/beginLetter/ingestCommandString/ingestReportString/titleShortFemale/titleFemale/gerundLabel/pawnLabel/stageName/shortDescription/customEffectDescriptions/endMessage/leaderTitle/pawnSingular/pawnsPlural/desc/recoveryMessage/chargeNoun/cooldownGerund/type/potentialExtraOutcomeDesc/labelNounPretty/headerTip/rejectInputMessage/spectatorGerund/spectatorsLabel/fuelLabel/formatString/useLabel/RMBLabel/permanentLabel/name/missingDesc/worshipRoomLabel/labelAbstract/fuelGizmoLabel/destroyedLabel/outOfFuelMessage/summary/ritualExpectedDesc/customSummary/meatLabel/labelForFullStatList/tooltip/gizmoLabel/onMapInstruction/letterTitle/textEnemy/destroyedOutLabel/beginLetterLabel/labelMale/groupName/gizmoDescription/names/arrivalTextEnemy/letterLabelEnemy/arrivedLetter/calledOffMessage/finishedMessage/approachingReportString/approachOrderString/expectedThingLabelTip/skillLabel/extraPredictedOutcomeDescriptions/modNameReadable/descriptionFuture/textWillArrive/arrivalTextFriendly/letterLabelFriendly/helpTextController/successfullyRemovedHediffMessage/textFriendly/eventLabel/textController/descOverride/shortDescOverride/content/discoveredLetterText/discoveredLetterTitle/beginLetterContinue/resourceLabel/message/overrideLabel/extraTooltip/offMessage/successMessage/effectDesc/letterInfoText/categoryLabel/groupLabel/battleStateLabel/customizationTitle/fixedName/noun/lockedReason/descriptionExtra/labelPrefix/labelMechanoids/ingestReportStringEat/failMessage/valueFormat/structureLabel/labelSocial/labelInBracketsExtraForHediff/ChooseDesc/ChooseLabel/ritualExplanation/resourceDescription/discoverLetterText/countdownLabel/inspectString/completedLetterText/completedLetterTitle/leaderDescription/formatStringUnfinalized/jobReportOverride/discoverLetterLabel/instantlyPermanentLabel/notifyMessage/onCooldownString/invalidTargetPawn/noAssignablePawnsDesc/reportText/statLabel/visualLabel/commandDescriptions/successMessageNoNegativeThought/tipLabelOverride/mainPartAllThreatsLabel/customChildDisallowMessage/ritualExpectedDescNoAdjective/loweredName/cancelLabel/texName/labelOverride/messageText/proficiencyAdjective/stuffAdjective/unit/labelTendedWell/labelTendedWellInner/labelSolidTendedWell/overrideTooltip/royalFavorLabel/extraReportString/spawnInBackstories/customLetterLabel/customLetterText/confirmationDialogText/tip/outcomeDescription/generalDescription/generalTitle/dialogue/activateDescString/activateLabelString/completedLetter/completedLetterLabel/guiLabelString/gizmoDesc/activatedMessageKey/appendString/gizmoDesc1/gizmoDesc2/gizmoLabel1/gizmoLabel".Split('/')
                ),
                FullListTranslationTags = new HashSet<string> { "rulesFiles", "rulesStrings", "pathList" },
                TranslationHandles = new List<string> { "*verbClass", "*compClass" },
                
                // 기존 NodeReplacement 복원
                NodeReplacement = new Dictionary<string, string>
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
                }
            };
        }

        public static string AutoDetectRimworldVersion()
        {
            try
            {
                var pathVersion = Path.Combine(Current.PathRimworld, "Version.txt");
                if (File.Exists(pathVersion))
                {
                    var context = File.ReadAllText(pathVersion).Trim();
                    var match = Regex.Match(context, Current.PatternVersion);
                    if (match.Success)
                    {
                        return match.Groups[0].Value;
                    }
                }
            }
            catch (Exception e)
            {
                Log.Err($"버전 자동 감지 실패: {e.Message}");
            }
            return Current.CurrentVersion;
        }
    }
}