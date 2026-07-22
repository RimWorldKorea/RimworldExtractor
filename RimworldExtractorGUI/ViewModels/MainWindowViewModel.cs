using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RimworldExtractorGUI.Utils;
using RimworldExtractorInternal;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorGUI.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    [ObservableProperty]
    private string _selectedModsText = "선택된 모드가 없습니다.\n\n===기본적인 사용방법===\n1) '1. 추출할 모드 선택'을 통해 번역할 모드를 선택하세요.\n선택) 번역할 모드를 우클릭 후, '이 모드와 관련된 모든 모드를 참조 모드로 선택'을 클릭하세요.\n2) '2. 번역 데이터 추출'을 통해 번역 텍스트를 추출하세요.\n3) 추출된 파일들을 번역하세요.\n5) 끝!\n\n팁) 옵션에서 출력 형식을 엑셀 파일(림왈도 형식)로 변경할 수 있습니다.";

    [ObservableProperty]
    private string _versionText = "버전 확인 중...";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ExtractCommand))]
    private bool _canExtract = false;

    public ModMetadata? SelectedMod { get; private set; }
    public List<ExtractableFolder>? SelectedFolders { get; private set; }
    public List<ModMetadata>? ReferenceMods { get; private set; }

    // View에 다이얼로그 호출을 요청하기 위한 이벤트들
    public event Func<Task>? RequestSelectMod;
    public event Func<Task>? RequestExtract;
    public event Func<Task>? RequestConvertXlsx;
    public event Func<Task>? RequestConvertXml;
    public event Func<Task>? RequestOpenSettings;
    public event Func<Task>? RequestOpenJpgPackager;
    public event Func<Task>? RequestOpenTranslationAnalyzer;

    public MainWindowViewModel()
    {
        CheckVersionAsync();
    }

    private void CheckVersionAsync()
    {
        Task.Run(() =>
        {
            try
            {
                var latest = GithubVersionCheker.GetLatest();
                var current = Program.VERSION;
                VersionText = latest == current
                    ? $"{current} 최신 버전입니다"
                    : $"{current} < {latest} 최신 버전 사용가능";
            }
            catch (Exception e)
            {
                Log.Wrn($"최신 버전 확인에 실패하였습니다: {e.Message}");
            }
        });
    }

    public void UpdateSelectedModInfo(ModMetadata mod, List<ExtractableFolder> folders, List<ModMetadata> refMods)
    {
        SelectedMod = mod;
        SelectedFolders = folders;
        ReferenceMods = refMods.Except(new[] { mod }).ToList();
        CanExtract = true;

        var text = $"선택된 모드: {SelectedMod.ModName}";
        if (ReferenceMods?.Count > 0)
        {
            var concatText = string.Join(", ", ReferenceMods.Select(x => x.ModName));
            var stripedText = concatText.Substring(0, Math.Min(concatText.Length, 200));
            if (concatText.Length > 200) stripedText += "...";
            text += $"\n참조로 선택된 모드: {stripedText}";
        }
        SelectedModsText = text;
    }

    [RelayCommand]
    private async Task SelectModAsync()
    {
        if (RequestSelectMod != null)
            await RequestSelectMod.Invoke();
    }

    [RelayCommand(CanExecute = nameof(CanExtract))]
    private async Task ExtractAsync()
    {
        if (RequestExtract != null)
            await RequestExtract.Invoke();
    }

    [RelayCommand]
    private async Task ConvertXlsxAsync()
    {
        if (RequestConvertXlsx != null)
            await RequestConvertXlsx.Invoke();
    }

    [RelayCommand]
    private async Task ConvertXmlAsync()
    {
        if (RequestConvertXml != null)
            await RequestConvertXml.Invoke();
    }

    [RelayCommand]
    private async Task OpenSettingsAsync()
    {
        if (RequestOpenSettings != null)
            await RequestOpenSettings.Invoke();
    }

    [RelayCommand]
    private async Task OpenJpgPackagerAsync()
    {
        if (RequestOpenJpgPackager != null)
            await RequestOpenJpgPackager.Invoke();
    }

    [RelayCommand]
    private async Task OpenTranslationAnalyzerAsync()
    {
        if (RequestOpenTranslationAnalyzer != null)
            await RequestOpenTranslationAnalyzer.Invoke();
    }

    [RelayCommand]
    private void OpenVersionUrl()
    {
        Process.Start(new ProcessStartInfo { FileName = GithubVersionCheker.LatestUrl, UseShellExecute = true });
    }

    [RelayCommand]
    private void OpenDiscussionUrl()
    {
        Process.Start(new ProcessStartInfo { FileName = GithubVersionCheker.DiscussionUrl, UseShellExecute = true });
    }
}