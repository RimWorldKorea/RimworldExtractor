using System.Text.RegularExpressions;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimExtractorCore;

namespace RimExtractorFace.ViewModels;

public partial class SettingsViewModel : ViewModelBase
{
    // --- UI 바인딩용 정적 배열 ---
    //TODO 여기서 따로 만들지 말고 코어모듈에 이미 있는거 가져다쓰기
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
    
    // 2차 언어용 목록 (없음 옵션 포함)
    public static string[] SecondaryLanguages { get; } = new[] { "None" }.Concat(Languages).ToArray();
    
    public static string[] ExtractionMethods { get; } = new[] { "스프레드시트 (.ods)", "XML 언어팩", "XML 언어팩 (주석 포함)" };
    public static string[] DuplicationPolicies { get; } = new[] { "작업중단", "덮어쓰기", "다른 이름으로 저장" };

    /// <summary>현재 설정의 캡쳐입니다.</summary>
    public Settings settings { get; }
    
    /// <summary>뷰의 버전 설정 사용자 입력을 담아두기 위한 속성입니다.</summary>
    [ObservableProperty]
    private string _inputVersion = string.Empty;
    
    public string gameVersion => settings.CurrentVersion;
    
    public string SelectedSecondaryLanguage
    {
        // 2차 언어 값이 비어있으면 "None"
        get => string.IsNullOrEmpty(settings.SecondaryLanguage) ? "None" : settings.SecondaryLanguage;
        set
        {
            settings.SecondaryLanguage = (value == "None") ? "" : value;
            OnPropertyChanged(nameof(SelectedSecondaryLanguage));
        }
    }

    // Enum ↔ ComboBox Index 어댑터 ---
    public int SelectedExtractionMethodIndex
    {
        get => (int)settings.Method;
        set
        {
            settings.Method = (ExportFileFormmat)value;
            OnPropertyChanged(nameof(SelectedExtractionMethodIndex));
        }
    }

    public int SelectedPolicyIndex
    {
        get => (int)settings.Policy;
        set
        {
            settings.Policy = (DuplicateFilePolicy)value;
            OnPropertyChanged(nameof(SelectedPolicyIndex));
        }
    }

    public event Action? RequestClose;
    public event Func<string, string, Task>? RequestShowAlert;

    public SettingsViewModel()
    {
        settings = SettingManager.Current.Clone();
        InputVersion = settings.CurrentVersion;
        RefreshAdapters();
    }

    private void RefreshAdapters()
    {
        OnPropertyChanged(nameof(settings));
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
            settings.PathRimworld = Path.GetDirectoryName(files[0].Path.LocalPath) ?? string.Empty;
            OnPropertyChanged(nameof(settings)); // 모델 데이터 변경 알림
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
            settings.PathWorkshop = folders[0].Path.LocalPath;
            OnPropertyChanged(nameof(settings));
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
            settings.PathBaseRefList = files[0].Path.LocalPath;
            OnPropertyChanged(nameof(settings));
        }
    }

    [RelayCommand]
    private void AutoDetectVersion()
    {
        SettingManager.UpdateGameVersion();
        OnPropertyChanged(nameof(gameVersion));
    }
    
    [RelayCommand]
    private void Apply()
    {
        ApplySettings(SettingManager.Current);
        
        Log.Good("일반 설정 적용 완료.", true);
    }
    
    [RelayCommand]
    private void Save()
    {
        ApplySettings(SettingManager.Current);
        ApplySettings(SettingManager.LastSaved, true);
        SettingManager.SaveToDisk();

        Log.Good("일반 설정 저장 완료.", true);
    }

    [RelayCommand]
    private void Close()
    {
        RequestClose?.Invoke();
    }
    
    private void ApplySettings(Settings target, bool SupressLogging = false)
    {
        target.PathRimworld = settings.PathRimworld;
        target.PathWorkshop = settings.PathWorkshop;
        target.PathBaseRefList = settings.PathBaseRefList;
        target.OriginalLanguage = settings.OriginalLanguage;
        target.SecondaryLanguage = settings.SecondaryLanguage;
        target.TranslationLanguage = settings.TranslationLanguage;
        target.CommentOriginal = settings.CommentOriginal;
        target.Policy = settings.Policy;
        target.Method = settings.Method;
        target.ExtractMayNotTranslate = settings.ExtractMayNotTranslate;
        
        string input = InputVersion.Trim();
        if (Regex.IsMatch(input, @"^\d+\.\d+$"))
        {
            target.CurrentVersion = input;
            settings.CurrentVersion = input;
        }
        else
        {
            InputVersion = target.CurrentVersion;
            string logText = "유효한 버전 형식이 아닙니다.";
            if (SupressLogging)
                Log.WrnOnce(logText, true);
            else
                Log.Wrn(logText, true);
        }
    }
}