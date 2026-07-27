using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.Input;
using RimExtractorCore;

namespace RimExtractorFace.ViewModels;

public partial class SettingsViewModel : ViewModelBase
{
    // --- UI 바인딩용 정적 배열 ---
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
    public static string[] ExtractionMethods { get; } = new[] { "엑셀 파일 (.xlsx)", "표준 언어팩 XML", "주석 포함 언어팩 XML" };
    public static string[] DuplicationPolicies { get; } = new[] { "중단", "덮어쓰기", "기존유지" };

    // --- 1. 모델 직접 바인딩 (XAML에서 {Binding Config.XXX} 사용) ---
    public Settings Config => ConfigManager.Current;

    // --- 2. Enum ↔ ComboBox Index 어댑터 ---
    public int SelectedExtractionMethodIndex
    {
        get => (int)Config.Method;
        set
        {
            Config.Method = (ExtractionMethod)value;
            OnPropertyChanged(nameof(SelectedExtractionMethodIndex));
        }
    }

    public int SelectedPolicyIndex
    {
        get => (int)Config.Policy;
        set
        {
            Config.Policy = (DuplicatesPolicy)value;
            OnPropertyChanged(nameof(SelectedPolicyIndex));
        }
    }

    // --- 3. Collection ↔ TextBox String 어댑터 ---
    public string ExtractableTagsText
    {
        get => string.Join('/', Config.ExtractableTags);
        set
        {
            Config.ExtractableTags = new HashSet<string>(RemoveSep(value).Split('/', StringSplitOptions.RemoveEmptyEntries));
            OnPropertyChanged(nameof(ExtractableTagsText));
        }
    }

    public string TranslationHandlesText
    {
        get => string.Join('/', Config.TranslationHandles);
        set
        {
            Config.TranslationHandles = new List<string>(RemoveSep(value).Split('/', StringSplitOptions.RemoveEmptyEntries));
            OnPropertyChanged(nameof(TranslationHandlesText));
        }
    }

    public string FullListTranslationTagsText
    {
        get => string.Join('/', Config.FullListTranslationTags);
        set
        {
            Config.FullListTranslationTags = new HashSet<string>(RemoveSep(value).Split('/', StringSplitOptions.RemoveEmptyEntries));
            OnPropertyChanged(nameof(FullListTranslationTagsText));
        }
    }

    public string NodeReplacementText
    {
        get => string.Join("/", Config.NodeReplacement.Select(x => $"{x.Key}|{x.Value}"));
        set
        {
            var replacements = RemoveSep(value).Split('/', StringSplitOptions.RemoveEmptyEntries);
            var dict = new Dictionary<string, string>();
            foreach (var r in replacements)
            {
                var token = r.Split('|');
                if (token.Length == 2) dict[token[0]] = token[1];
            }
            Config.NodeReplacement = dict;
            OnPropertyChanged(nameof(NodeReplacementText));
        }
    }

    public event Action? RequestClose;
    public event Func<string, string, Task>? RequestShowAlert;

    public SettingsViewModel()
    {
        ConfigManager.Load();
        RefreshAdapters();
    }

    private void RefreshAdapters()
    {
        OnPropertyChanged(nameof(Config));
        OnPropertyChanged(nameof(SelectedExtractionMethodIndex));
        OnPropertyChanged(nameof(SelectedPolicyIndex));
        OnPropertyChanged(nameof(ExtractableTagsText));
        OnPropertyChanged(nameof(TranslationHandlesText));
        OnPropertyChanged(nameof(FullListTranslationTagsText));
        OnPropertyChanged(nameof(NodeReplacementText));
    }

    private static string RemoveSep(string s) => s?.Replace(" ", "").Replace("\r", "").Replace("\n", "") ?? "";

    // ----------------------------------------------------
    // 커맨드 영역
    // ----------------------------------------------------

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
            Config.PathRimworld = System.IO.Path.GetDirectoryName(files[0].Path.LocalPath) ?? string.Empty;
            OnPropertyChanged(nameof(Config)); // 모델 데이터 변경 알림
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
            Config.PathWorkshop = folders[0].Path.LocalPath;
            OnPropertyChanged(nameof(Config));
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
            Config.PathBaseRefList = files[0].Path.LocalPath;
            OnPropertyChanged(nameof(Config));
        }
    }

    [RelayCommand]
    private void AutoDetectVersion()
    {
        Config.CurrentVersion = ConfigManager.AutoDetectRimworldVersion();
        OnPropertyChanged(nameof(Config));
    }

    [RelayCommand]
    private void SaveAndClose()
    {
        ConfigManager.Save();
        RequestClose?.Invoke();
    }

    [RelayCommand]
    private void Cancel()
    {
        ConfigManager.Load(); // 변경 사항 롤백
        RefreshAdapters();
        RequestClose?.Invoke();
    }

    [RelayCommand]
    private void Reset()
    {
        ConfigManager.InitDefault(); // 기본값으로 덮어쓰기
        RefreshAdapters();
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