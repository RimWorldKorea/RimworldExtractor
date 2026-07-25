using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimworldExtractorGUI.Services;
using RimworldExtractorInternal.Core;

namespace RimworldExtractorGUI.ViewModels;

public partial class SettingsViewModel : ViewModelBase
{
    private readonly IPrefabSettingsService _settingsService;

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
        "엑셀 파일 (.xlsx)",
        "표준 언어팩 XML",
        "주석 포함 언어팩 XML"
    };

    public static string[] DuplicationPolicies { get; } = new[]
    {
        "중단", "덮어쓰기", "기존유지"
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

    public SettingsViewModel(IPrefabSettingsService settingsService)
    {
        _settingsService = settingsService;
        _settingsService.Load();
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
            Title = "RimWorldWin64.exe 파일 선택",
            AllowMultiple = false,
            FileTypeFilter = new[] { new FilePickerFileType("실행 파일") { Patterns = new[] { "RimWorldWin64.exe" } } }
        });
        if (files.Count > 0)
        {
            PathRimworld = System.IO.Path.GetDirectoryName(files[0].Path.LocalPath) ?? string.Empty;
        }
    }

    [RelayCommand]
    private async Task SelectPathWorkshopAsync(IStorageProvider storageProvider)
    {
        var folders = await storageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "창작마당 폴더 => Steam\\steamapps\\workshop\\content\\294100",
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
            Title = "기준 모드 파일 선택",
            AllowMultiple = false,
            FileTypeFilter = new[] { new FilePickerFileType("기준 모드 파일") { Patterns = new[] { "*.refMods" } } }
        });
        if (files.Count > 0)
        {
            PathBaseRefList = files[0].Path.LocalPath;
        }
    }

    [RelayCommand]
    private void AutoDetectVersion()
    {
        CurrentVersion = _settingsService.AutoDetectVersion();
    }

    [RelayCommand]
    private void SaveAndClose()
    {
        ToPrefabs();
        _settingsService.Save();
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
        _settingsService.Reset();
        FromPrefabs();
    }

    #region Help Commands
    [RelayCommand]
    private async Task ShowHelp1Async()
    {
        if (RequestShowAlert != null)
            await RequestShowAlert.Invoke("도움말", "추출할 태그를 추가/삭제합니다. '/' 구분자를 사용합니다.");
    }

    [RelayCommand]
    private async Task ShowHelp2Async()
    {
        if (RequestShowAlert != null)
            await RequestShowAlert.Invoke("도움말", "Translation Handle은 번역할 태그가 'li'로 되어있을 때 사용합니다. https://ludeon.com/forums/index.php?topic=41942.0 참고\n" +
                "해당 태그의 자식 태그의 이름을 Translation Handle에 추가해주면 됩니다.\n" +
                "예시) verbs.2.label => verbs.Verb_Shoot.label\n'/' 구분자를 사용합니다.\n" +
                "Translation Handle의 자식 태그가 Type일 경우 해당 Type 앞에 '*'를 붙여줍니다.");
    }

    [RelayCommand]
    private async Task ShowHelp3Async()
    {
        if (RequestShowAlert != null)
            await RequestShowAlert.Invoke("도움말", "태그를 치환해줍니다. " +
                "'(Def 타입)+(태그명)|(Def 타입)+(바꿀 태그명)' 과 같이 적어주며 defName 노드 이전 노드 명은 모두 '/'로 생략해야 합니다. " +
                "\n바꿀 태그명에 'li' 태그가 포함되어선 안됩니다.");
    }

    [RelayCommand]
    private async Task ShowHelp4Async()
    {
        if (RequestShowAlert != null)
            await RequestShowAlert.Invoke("도움말", "Full-list Translation 태그를 설정합니다. https://ludeon.com/forums/index.php?topic=41942.0 참고\n '/' 구분자를 사용합니다.");
    }
    #endregion
}