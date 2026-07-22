using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimworldExtractorInternal;

namespace RimworldExtractorGUI.ViewModels;

public partial class SettingsViewModel : ViewModelBase
{
    public static string[] Languages { get; } = new[]
    {
        "English", "Korean (한국어)", "Catalan (Català)", "ChineseSimplified (简体中文)", "ChineseTraditional (繁體中文)",
        "Czech (Čeština)", "Danish (Dansk)", "Dutch (Nederlands)", "Estonian (Eesti)",
        "Finnish (Suomi)", "French (Français)", "German (Deutsch)", "Greek (Ελληνικά)",
        "Hungarian (Magyar)", "Italian (Italiano)", "Japanese (日本語)", "Norwegian (Norsk Bokmål)",
        "Polish (Polski)", "Portuguese (Português)", "PortugueseBrazilian (Português Brasileiro)",
        "Romanian (Română)", "Russian (Русский)", "Slovak (Slovenčina)", "Spanish (Español(Castellano))",
        "SpanishLatin (Español(Latinoamérica))", "Swedish (Svenska)", "Turkish (Türkçe)",
        "Ukrainian (Українська)"
    };

    public static string[] ExtractionMethods { get; } = new[]
    {
        "번역 작업에 쓰일 엑셀(.xlsx) 파일",
        "배포 가능한 XML 파일",
        "배포 가능한 XML 파일(주석 포함)"
    };

    public static string[] DuplicationPolicies { get; } = new[]
    {
        "멈추고 묻기", "덮어씌우기", "건너뛰기"
    };

    [ObservableProperty] private string _pathRimworld = string.Empty;
    [ObservableProperty] private string _pathWorkshop = string.Empty;
    [ObservableProperty] private string _patternVersion = string.Empty;
    [ObservableProperty] private string _currentVersion = string.Empty;

    [ObservableProperty] private string _originalLanguage = "English";
    [ObservableProperty] private string _translationLanguage = "Korean (한국어)";
    [ObservableProperty] private int _selectedExtractionMethodIndex = 0;
    [ObservableProperty] private int _selectedPolicyIndex = 0;
    [ObservableProperty] private string _pathBaseRefList = string.Empty;

    [ObservableProperty] private string _extractableTags = string.Empty;
    [ObservableProperty] private string _translationHandles = string.Empty;
    [ObservableProperty] private string _nodeReplacement = string.Empty;
    [ObservableProperty] private string _fullListTranslationTags = string.Empty;
    [ObservableProperty] private bool _enableTkey = false;

    public event Action? RequestClose;
    public event Func<string, string, Task>? RequestShowAlert;

    public SettingsViewModel()
    {
        if (File.Exists("Prefabs.dat"))
        {
            Prefabs.Load();
        }
        else
        {
            Prefabs.Save();
        }

        FromPrefabs();
    }

    public void FromPrefabs()
    {
        EnableTkey = Prefabs.EnableTkey;

        PathRimworld = Prefabs.PathRimworld;
        PathWorkshop = Prefabs.PathWorkshop;
        PatternVersion = Prefabs.PatternVersion;
        CurrentVersion = Prefabs.CurrentVersion;

        OriginalLanguage = Prefabs.OriginalLanguage;
        TranslationLanguage = Prefabs.TranslationLanguage;
        SelectedExtractionMethodIndex = (int)Prefabs.Method;
        SelectedPolicyIndex = (int)Prefabs.Policy;
        PathBaseRefList = Prefabs.PathBaseRefList;

        ExtractableTags = string.Join('/', Prefabs.ExtractableTags);
        TranslationHandles = string.Join('/', Prefabs.TranslationHandles);
        NodeReplacement = string.Join("/", Prefabs.NodeReplacement.Select(x => $"{x.Key}|{x.Value}"));
        FullListTranslationTags = string.Join('/', Prefabs.FullListTranslationTags);
    }

    public void ToPrefabs()
    {
        Prefabs.EnableTkey = EnableTkey;

        Prefabs.PathRimworld = PathRimworld ?? string.Empty;
        Prefabs.PathWorkshop = PathWorkshop ?? string.Empty;
        Prefabs.PatternVersion = PatternVersion ?? string.Empty;
        Prefabs.CurrentVersion = CurrentVersion ?? string.Empty;

        Prefabs.OriginalLanguage = OriginalLanguage ?? "English";
        Prefabs.TranslationLanguage = TranslationLanguage ?? "Korean (한국어)";

        if (SelectedExtractionMethodIndex >= 0)
            Prefabs.Method = Enum.GetValues<Prefabs.ExtractionMethod>()[SelectedExtractionMethodIndex];

        if (SelectedPolicyIndex >= 0)
            Prefabs.Policy = Enum.GetValues<Prefabs.DuplicatesPolicy>()[SelectedPolicyIndex];

        Prefabs.PathBaseRefList = PathBaseRefList ?? string.Empty;

        Prefabs.ExtractableTags = new HashSet<string>(RemoveSep(ExtractableTags ?? "").Split('/', StringSplitOptions.RemoveEmptyEntries));
        Prefabs.TranslationHandles = new List<string>(RemoveSep(TranslationHandles ?? "").Split('/', StringSplitOptions.RemoveEmptyEntries));

        var replacements = RemoveSep(NodeReplacement ?? "").Split('/', StringSplitOptions.RemoveEmptyEntries);
        var dict = new Dictionary<string, string>();
        foreach (var r in replacements)
        {
            var token = r.Split('|');
            if (token.Length == 2)
            {
                dict[token[0]] = token[1];
            }
        }
        Prefabs.NodeReplacement = dict;

        Prefabs.FullListTranslationTags = new HashSet<string>(RemoveSep(FullListTranslationTags ?? "").Split('/', StringSplitOptions.RemoveEmptyEntries));
    }

    private static string RemoveSep(string s) => s.Replace(" ", "").Replace("\r", "").Replace("\n", "");

    [RelayCommand]
    private async Task SelectPathRimworldAsync(IStorageProvider storageProvider)
    {
        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "RimWorldWin64.exe를 지정해주세요",
            AllowMultiple = false,
            FileTypeFilter = new[] { new FilePickerFileType("림월드 실행 파일") { Patterns = new[] { "RimWorldWin64.exe" } } }
        });

        if (files.Count > 0)
        {
            PathRimworld = Path.GetDirectoryName(files[0].Path.LocalPath) ?? string.Empty;
        }
    }

    [RelayCommand]
    private async Task SelectPathWorkshopAsync(IStorageProvider storageProvider)
    {
        var folders = await storageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "림월드 창작마당 경로를 지정해주세요 => Steam\\steamapps\\workshop\\content\\294100",
            AllowMultiple = false
        });

        if (folders.Count > 0)
        {
            PathWorkshop = folders[0].Path.LocalPath;
        }
    }

    [RelayCommand]
    private async Task SelectBaseRefListAsync(IStorageProvider storageProvider)
    {
        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "선택한 파일로부터 참조 모드의 목록을 불러옵니다.",
            AllowMultiple = false,
            FileTypeFilter = new[] { new FilePickerFileType("참조 모드 리스트 파일") { Patterns = new[] { "*.refMods" } } }
        });

        if (files.Count > 0)
        {
            PathBaseRefList = files[0].Path.LocalPath;
        }
    }

    [RelayCommand]
    private void AutoDetectVersion()
    {
        CurrentVersion = Prefabs.AutoDetectRimworldVersion();
    }

    [RelayCommand]
    private void SaveAndClose()
    {
        ToPrefabs();
        Prefabs.Save();
        RequestClose?.Invoke();
    }

    [RelayCommand]
    private void Cancel()
    {
        RequestClose?.Invoke();
    }

    [RelayCommand]
    private void Reset()
    {
        Prefabs.Init();
        Prefabs.Save();
        FromPrefabs();
    }

    #region Help Commands

    [RelayCommand]
    private async Task ShowHelp1Async()
    {
        if (RequestShowAlert != null)
            await RequestShowAlert.Invoke("도움말", "추출해야 하는 노드의 태그 목록을 정의합니다. '/' 문자로 구분하여 공백 없이 입력합니다. 특별한 일이 없는 이상 기본 상태로 두세요.");
    }

    [RelayCommand]
    private async Task ShowHelp2Async()
    {
        if (RequestShowAlert != null)
            await RequestShowAlert.Invoke("도움말", "Translation Handle은 노드의 이름이 'li'인 리스트 노드를 추출할 때 특정 태그의 값을 리스트 번호 대신 사용하는 추출 방법입니다. https://ludeon.com/forums/index.php?topic=41942.0 참고\n" +
                "해당 노드에 Translation Handle 태그와 일치하는 노드가 있으면 그 노드의 값으로 리스트 노드의 이름을 결정합니다.\n" +
                "예) verbs.2.label => verbs.Verb_Shoot.label\n'/' 문자로 구분하여 공백 없이 입력합니다.\n" +
                "Translation Handle의 추출 방식은 그 태그의 타입이 Type 타입인지, 그 외인지에 따라 다릅니다. 따라서 Type 타입인 경우 앞에 접두어 '*'를 붙입니다.");
    }

    [RelayCommand]
    private async Task ShowHelp3Async()
    {
        if (RequestShowAlert != null)
            await RequestShowAlert.Invoke("도움말", "일부 노드는 추출했을 때의 노드와 번역을 적용할 때의 노드가 다른데, 노드 대체 기능은 그러한 경우에 사용됩니다. " +
                "'(Def 타입)+(원본 노드)|(Def 타입)+(대체 노드)의 형식으로 입력하며, 이때 defName 부분은 생략하여 입력합니다. 여러 개인 경우 '/' 문자로 구분하여 공백 없이 입력합니다. " +
                "\n주의: 아직 'li' 노드가 포함된 경우는 지원하지 않습니다.");
    }

    [RelayCommand]
    private async Task ShowHelp4Async()
    {
        if (RequestShowAlert != null)
            await RequestShowAlert.Invoke("도움말", "Full-list Translation은 일부 경우에만 사용되는 리스트 노드 저장 방식입니다. https://ludeon.com/forums/index.php?topic=41942.0 참고\n '/' 문자로 구분하여 공백 없이 입력합니다.");
    }

    #endregion
}