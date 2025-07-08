using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Reflection;
using System.Text.RegularExpressions;

namespace BatchConvertToRVZ;

/// <summary>
/// Service for checking for application updates on GitHub.
/// </summary>
public sealed class UpdateService : IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly string _githubApiUrl;

    /// <summary>
    /// Initializes a new instance of the <see cref="UpdateService"/> class.
    /// </summary>
    /// <param name="githubApiUrl">The URL to the GitHub API for the latest release.</param>
    public UpdateService(string githubApiUrl)
    {
        _githubApiUrl = githubApiUrl;
        _httpClient = new HttpClient();
        // GitHub API requires a User-Agent header.
        _httpClient.DefaultRequestHeaders.UserAgent.Add(
            new ProductInfoHeaderValue("BatchConvertToRVZ", GetCurrentVersion().ToString()));
    }

    /// <summary>
    /// Checks for a new version of the application.
    /// </summary>
    /// <returns>
    /// A tuple containing a boolean indicating if an update is available
    /// and the <see cref="GitHubRelease"/> object if an update is found.
    /// </returns>
    public async Task<(bool IsUpdateAvailable, GitHubRelease? LatestRelease)> CheckForUpdatesAsync()
    {
        try
        {
            var latestRelease = await _httpClient.GetFromJsonAsync<GitHubRelease>(_githubApiUrl);

            if (latestRelease == null || latestRelease.Draft || latestRelease.Prerelease)
            {
                return (false, null);
            }

            var currentVersion = GetCurrentVersion();
            var latestVersion = ParseVersionFromTag(latestRelease.TagName);

            if (latestVersion != null && latestVersion > currentVersion)
            {
                return (true, latestRelease);
            }
        }
        catch (Exception)
        {
            // Could be a network error, JSON parsing error, etc. Silently fail for automatic checks.
            return (false, null);
        }

        return (false, null);
    }

    /// <summary>
    /// Gets the current version of the executing assembly.
    /// </summary>
    private static Version GetCurrentVersion()
    {
        return Assembly.GetExecutingAssembly().GetName().Version ?? new Version("0.0.0.0");
    }

    /// <summary>
    /// Parses a <see cref="Version"/> object from a GitHub tag name.
    /// </summary>
    /// <param name="tagName">The tag name (e.g., "v1.2.3" or "release-1.2.3").</param>
    /// <returns>A <see cref="Version"/> object or null if parsing fails.</returns>
    private static Version? ParseVersionFromTag(string tagName)
    {
        // Remove any non-numeric/non-dot prefixes (like 'v' or 'release-').
        var versionString = Regex.Replace(tagName, "^[^0-9]+", "");
        return Version.TryParse(versionString, out var version) ? version : null;
    }

    public void Dispose()
    {
        _httpClient.Dispose();
    }
}
