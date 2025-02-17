﻿using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace RimworldExtractorInternal
{
    public static class Prefabs
    {
        /// <summary>
        /// Prefabs.dat의 호환성을 위해 존재합니다. Prefabs의 필드가 수정되었을 때 이 숫자를 1 증가시켜,
        /// Prefabs.dat에 저장된 숫자가 이와 다르다면, 해당 데이터를 읽지 않도록 합니다.
        /// </summary>
        private static readonly string Version = "8";

        // 임시 기능 TODO: REMOVE THIS AFTER
        public static bool EnableTkey = false;

        public static string PathRimworld = string.Empty;
        public static string PathWorkshop = string.Empty;
        public static string PathBaseRefList = "";
        public static string CurrentVersion = string.Empty;
        public static string PatternVersion = string.Empty;
        public static string PatternVersionWithV = string.Empty;
        public static string OriginalLanguage = string.Empty;
        public static string TranslationLanguage = string.Empty;
        public static bool CommentOriginal = false;

        public static HashSet<string> ExtractableTags = new();

        public static HashSet<string> FullListTranslationTags = new();

        public static Dictionary<string, string> NodeReplacement = new();

        /// <summary>
        /// TranslationHandle ordered by Priority
        /// </summary>
        public static List<string> TranslationHandles = new();

        public static DuplicatesPolicy Policy = default;
        public static ExtractionMethod Method = default;


        public static Action<XLWorkbook, string>? StopCallbackXlsx = null;
        public static Action<XmlDocument, string>? StopCallbackXml = null; 
        public static Action<IEnumerable<string>, string>? StopCallbackTxt = null;


        public const string ExtensionKeyExtraCommentTranslated = "ExtraCommentTranslated";

        public static void Init()
        {
            // 임시 기능 TODO: REMOVE THIS AFTER
            EnableTkey = false;
            PathRimworld = "C:\\Program Files (x86)\\Steam\\steamapps\\common\\RimWorld";
            PathWorkshop = "C:\\Program Files (x86)\\Steam\\steamapps\\workshop\\content\\294100";
            PathBaseRefList = "";
            CurrentVersion = "1.5";
            PatternVersion = @"^[1]\.\d+";
            PatternVersionWithV = @"^v[1]\.\d+";
            OriginalLanguage = "English";
            TranslationLanguage = "Korean (한국어)";
            CommentOriginal = false;
            ExtractableTags = new(
                "label/rulesStrings/description/baseDesc/title/titleShort/customLabel/symbol/jobString/reportString/labelNoun/slateRef/verb/gerund/adjective/member/tips/ideoName/thoughtStageDescriptions/jobReportString/theme/labelShortAdj/labelPlural/letterText/deathMessage/labelShort/letterLabel/helpText/text/baseInspectLine/labelFemale/descriptionShort/beginLetter/ingestCommandString/ingestReportString/titleShortFemale/titleFemale/gerundLabel/pawnLabel/stageName/shortDescription/customEffectDescriptions/endMessage/leaderTitle/pawnSingular/pawnsPlural/desc/recoveryMessage/chargeNoun/cooldownGerund/type/potentialExtraOutcomeDesc/labelNounPretty/headerTip/rejectInputMessage/spectatorGerund/spectatorsLabel/fuelLabel/formatString/useLabel/RMBLabel/permanentLabel/name/missingDesc/worshipRoomLabel/labelAbstract/fuelGizmoLabel/destroyedLabel/outOfFuelMessage/summary/ritualExpectedDesc/customSummary/meatLabel/labelForFullStatList/tooltip/gizmoLabel/onMapInstruction/letterTitle/textEnemy/destroyedOutLabel/beginLetterLabel/labelMale/groupName/gizmoDescription/names/arrivalTextEnemy/letterLabelEnemy/arrivedLetter/calledOffMessage/finishedMessage/approachingReportString/approachOrderString/expectedThingLabelTip/skillLabel/extraPredictedOutcomeDescriptions/modNameReadable/descriptionFuture/textWillArrive/arrivalTextFriendly/letterLabelFriendly/helpTextController/successfullyRemovedHediffMessage/textFriendly/eventLabel/textController/descOverride/shortDescOverride/content/discoveredLetterText/discoveredLetterTitle/beginLetterContinue/resourceLabel/message/overrideLabel/extraTooltip/offMessage/successMessage/effectDesc/letterInfoText/categoryLabel/groupLabel/battleStateLabel/customizationTitle/fixedName/noun/lockedReason/descriptionExtra/labelPrefix/labelMechanoids/ingestReportStringEat/failMessage/valueFormat/structureLabel/labelSocial/labelInBracketsExtraForHediff/ChooseDesc/ChooseLabel/ritualExplanation/resourceDescription/discoverLetterText/countdownLabel/inspectString/completedLetterText/completedLetterTitle/leaderDescription/formatStringUnfinalized/jobReportOverride/discoverLetterLabel/instantlyPermanentLabel/notifyMessage/onCooldownString/invalidTargetPawn/noAssignablePawnsDesc/reportText/statLabel/visualLabel/commandDescriptions/successMessageNoNegativeThought/tipLabelOverride/mainPartAllThreatsLabel/customChildDisallowMessage/ritualExpectedDescNoAdjective/loweredName/cancelLabel/texName/labelOverride/messageText/proficiencyAdjective/stuffAdjective/unit/labelTendedWell/labelTendedWellInner/labelSolidTendedWell/overrideTooltip/royalFavorLabel/extraReportString/spawnInBackstories/customLetterLabel/customLetterText/confirmationDialogText/tip/outcomeDescription/generalDescription/generalTitle/dialogue/activateDescString/activateLabelString/completedLetter/completedLetterLabel/guiLabelString"
                    .Split('/'));
            FullListTranslationTags = new()
            {
                "rulesFiles", "rulesStrings", "pathList"
            };
            NodeReplacement = new Dictionary<string, string>(
                "ScenarioDef+label|ScenarioDef+scenario.name/ScenarioDef+description|ScenarioDef+scenario.description/CombatExtended.AmmoDef+*|ThingDef+*/VFECore.ExpandableProjectileDef+*|ThingDef+*/AbilityUser.ProjectileDef_AbilityLaser+*|ThingDef+*/AbilityUser.ProjectileDef_Ability+*|ThingDef+*/NewRatkin.CustomThingDef+*|ThingDef+*/AlienRace.AlienBackstoryDef+*|BackstoryDef+*/RatkinGeneExpanded.FactionDefExtended+*|FactionDef+*/RatkinGeneExpanded.ThingDefExtended+*|ThingDef+*/AlienRace.ThingDef_AlienRace+*|ThingDef+*/Rimlaser.Building_LaserGunDef+*|ThingDef+*/Rimlaser.LaserBeamDef+*|ThingDef+*/Rimlaser.LaserGunDef+*|ThingDef+*/Rimlaser.SpinningLaserGunDef+*|ThingDef+*/JecsTools.BackstoryDef+baseDesc|JescTools.BackstoryDef+description/AnestheticGunMod2.AnestheticBulletDef+*|ThingDef+*/BackstoryDef+baseDesc|BackstoryDef+description/DubsBadHygiene.WashingJobDef+*|JobDef+*/DubsBadHygiene.Needy+*|NeedDef+*/VarietyMatters.FoodVariety_NeedDef+*|NeedDef+*/Kiiro.StorytellerDef_Custom+*|StorytellerDef+*/Vehicles.SkinDef+*|Vehicles.PatternDef+*/Vehicles.AntiAircraftDef+*|WorldObjectDef+*/Vehicles.AirdropDef+*|ThingDef+*"
                    .Split("/").Select(
                        x =>
                        {
                            var tokens = x.Split('|');
                            return new KeyValuePair<string, string>(tokens[0].Trim(), tokens[1].Trim());
                        }));
            TranslationHandles = new()
            {
                // "label", // 200
                // "customLabel", "name", "def",
                // "inSignal", "labelMale", // 100
                // "labelFemale", 
                "*verbClass", "*compClass", 
                // "hediff"
            };
            Policy = DuplicatesPolicy.Overwrite;
            Method = ExtractionMethod.Languages;
        }

        public static void Save(string fileName = "Prefabs.dat")
        {
            List<string> lines = new List<string>
            {
                "DO NOT EDIT THIS MANUALLY",
                Version,
                EnableTkey.ToString(), // TODO: REMOVE THIS AFTER
                PathRimworld,
                PathWorkshop,
                PathBaseRefList,
                CurrentVersion,
                PatternVersion,
                PatternVersionWithV,
                OriginalLanguage,
                TranslationLanguage,
                CommentOriginal.ToString(),
                string.Join('/', ExtractableTags),
                string.Join('/', FullListTranslationTags),
                string.Join('/', NodeReplacement.Select(x => $"{x.Key}|{x.Value}")),
                string.Join('/', TranslationHandles),
                Policy.ToString(),
                Method.ToString()
            };
            File.WriteAllLines(fileName, lines);
        }


        /// <exception cref="SerializationException">Version 필드의 값이 달라서 생기는 에러</exception>
        public static void Load(string fileName = "Prefabs.dat")
        {
            var lines = File.ReadAllLines(fileName);
            var idx = 1;
            if (Version != lines[idx++])
            {
                throw new SerializationException($"wrong version of {fileName}");
            }

            EnableTkey = bool.Parse(lines[idx++]); // TODO: REMOVE THIS AFTER
            PathRimworld = lines[idx++];
            PathWorkshop = lines[idx++];
            PathBaseRefList = lines[idx++];
            CurrentVersion = lines[idx++];
            PatternVersion = lines[idx++];
            PatternVersionWithV = lines[idx++];
            OriginalLanguage = lines[idx++];
            TranslationLanguage = lines[idx++];
            CommentOriginal = bool.Parse(lines[idx++]);
            ExtractableTags = new(lines[idx++].Split('/'));
            FullListTranslationTags = new(lines[idx++].Split("/"));
            NodeReplacement = new(lines[idx++].Split('/').Select(x =>
            {
                var splited = x.Split('|');
                return new KeyValuePair<string, string>(splited[0], splited[1]);
            }));
            TranslationHandles = new(lines[idx++].Split('/'));
            Policy = Enum.Parse<DuplicatesPolicy>(lines[idx++]);
            Method = Enum.Parse<ExtractionMethod>(lines[idx++]);
        }

        public static string AutoDetectRimworldVersion()
        {
            try
            {
                var pathVersion = Path.Combine(PathRimworld, "Version.txt");
                var context = File.ReadAllText(pathVersion).Trim();
                var match = Regex.Match(context, PatternVersion);
                if (match.Success)
                {
                    var version = match.Groups[0].Value;
                    return version;
                }
            }
            catch (Exception e)
            {
                Log.Err($"버전 자동 감지 중 에러 발생 {e.Message}");
            }

            return CurrentVersion;
        }

        public enum DuplicatesPolicy
        {
            Stop = 0,
            Overwrite,
            KeepOriginal
        }

        public enum ExtractionMethod
        {
            Excel = 0, Languages, LanguagesWithComments
        }
    }
}
