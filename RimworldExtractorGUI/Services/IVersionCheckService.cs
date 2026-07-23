namespace RimworldExtractorGUI.Services;

public interface IVersionCheckService
{
    /// <summary>
    /// 현재 애플리케이션의 버전을 가져옵니다.
    /// </summary>
    string CurrentVersion { get; }

    string ReleasesUrl { get; }
    string LatestUrl { get; }
    string IssueUrl { get; }
    string DiscordUrl { get; }

    /// <summary>
    /// 서버(GitHub)에서 최신 릴리즈 버전을 비동기적으로 가져옵니다.
    /// </summary>
    Task<string> GetLatestVersionAsync();
}