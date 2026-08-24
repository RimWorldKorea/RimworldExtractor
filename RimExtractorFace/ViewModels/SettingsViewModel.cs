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
    
    // [추가됨] 2차 언어용 목록 (없음 옵션 포함)
    public static string[] SecondaryLanguages { get; } = new[] { "None" }.Concat(Languages).ToArray();
    
    public static string[] ExtractionMethods { get; } = new[] { "스프레드 시트 (.ods)", "표준 언어팩 XML", "주석 포함 언어팩 XML" };
    public static string[] DuplicationPolicies { get; } = new[] { "중단", "덮어쓰기", "기존유지" };

    // --- 1. 모델 직접 바인딩 (XAML에서 {Binding Config.XXX} 사용) ---
    public Settings Config => SettingManager.Current;
    
    // [추가됨] 2차 언어 값이 비어있을 때 "None"으로 매핑해주는 Wrapper
    public string SelectedSecondaryLanguage
    {
        get => string.IsNullOrEmpty(Config.SecondaryLanguage) ? "None" : Config.SecondaryLanguage;
        set
        {
            Config.SecondaryLanguage = (value == "None") ? "" : value;
            OnPropertyChanged(nameof(SelectedSecondaryLanguage));
        }
    }

    // --- 2. Enum ↔ ComboBox Index 어댑터 ---
    public int SelectedExtractionMethodIndex
    {
        get => (int)Config.Method;
        set
        {
            Config.Method = (ExportFileFormmat)value;
            OnPropertyChanged(nameof(SelectedExtractionMethodIndex));
        }
    }

    public int SelectedPolicyIndex
    {
        get => (int)Config.Policy;
        set
        {
            Config.Policy = (DuplicateFilePolicy)value;
            OnPropertyChanged(nameof(SelectedPolicyIndex));
        }
    }

    public event Action? RequestClose;
    public event Func<string, string, Task>? RequestShowAlert;

    public SettingsViewModel()
    {
        SettingManager.Load();
        RefreshAdapters();
    }

    private void RefreshAdapters()
    {
        OnPropertyChanged(nameof(Config));
        OnPropertyChanged(nameof(SelectedExtractionMethodIndex));
        OnPropertyChanged(nameof(SelectedPolicyIndex));
        OnPropertyChanged(nameof(SelectedSecondaryLanguage));
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
        Config.CurrentVersion = SettingManager.AutoDetectRimworldVersion();
        OnPropertyChanged(nameof(Config));
    }

    [RelayCommand]
    private void SaveAndClose()
    {
        SettingManager.Save();
        RequestClose?.Invoke();
    }

    [RelayCommand]
    private void Cancel()
    {
        SettingManager.Load(); // 변경 사항 롤백
        RefreshAdapters();
        RequestClose?.Invoke();
    }

    [RelayCommand]
    private void Reset()
    {
        SettingManager.InitDefault(); // 기본값으로 덮어쓰기
        RefreshAdapters();
    }
}