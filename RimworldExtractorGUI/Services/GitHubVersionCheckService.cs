using System.Net;

namespace RimworldExtractorGUI.Services;

public class GitHubVersionCheckService : IVersionCheckService
{
    // HttpClient는 소켓 고갈(Socket Exhaustion) 방지를 위해 애플리케이션 수명 주기 동안 재사용하는 것이 권장됩니다.
    private static readonly HttpClient _httpClient = new HttpClient(new HttpClientHandler
    {
        AllowAutoRedirect = false 
    });

    public string CurrentVersion => Program.VERSION;

    public string ReleasesUrl => "https://github.com/csh1668/RimworldExtractor/releases";
    
    // Path.Combine 대신 URL에 안전한 문자열 보간을 사용합니다.
    public string LatestUrl => $"{ReleasesUrl}/latest";
    
    public string IssueUrl => "https://github.com/csh1668/RimworldExtractor/issues/new/choose";
    
    public string DiscordUrl => "https://discord.gg/5FdkKj2XUe";

    public async Task<string> GetLatestVersionAsync()
    {
        // .Result 대신 async/await를 사용하여 UI 스레드 블로킹을 방지합니다.
        var response = await _httpClient.GetAsync(LatestUrl);

        if (response.StatusCode is HttpStatusCode.Redirect or HttpStatusCode.MovedPermanently or HttpStatusCode.Found)
        {
            var redirectedUrl = response.Headers.Location;
            return redirectedUrl?.AbsolutePath.Split('/').Last() ?? throw new WebException("RedirectedUrl was null.");
        }

        throw new WebException($"HttpClient got unexpected response: {response.StatusCode}");
    }
}