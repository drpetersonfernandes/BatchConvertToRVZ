using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Reflection;
using System.Text.RegularExpressions;
using BatchConvertToRVZ.models;

namespace BatchConvertToRVZ.services;

/// <summary>
/// Service for checking for application updates on GitHub.
/// </summary>
public partial class UpdateService : IDisposable
{
    // Shared static HttpClient instance to prevent socket exhaustion
    // SocketsHttpHandler with PooledConnectionLifetime ensures proper connection pooling
    private static readonly SocketsHttpHandler SharedHandler = new()
    {
        PooledConnectionLifetime = TimeSpan.FromMinutes(2)
    };

    private readonly HttpClient _httpClient;
    private readonly string _githubApiUrl;

    /// <summary>
    /// Initializes a new instance of the <see cref="UpdateService"/> class.
    /// </summary>
    /// <param name="githubApiUrl">The URL to the GitHub API for the latest release.</param>
    public UpdateService(string githubApiUrl)
    {
        _githubApiUrl = githubApiUrl;

        // Create a new HttpClient instance that shares the static handler
        // This allows per-instance headers while sharing the connection pool
        _httpClient = new HttpClient(SharedHandler, false);

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
        // Extract only the numeric version components (e.g., "1.2.3" or "1.2.3.4")
        // This handles prefixes like "v" or "release-" and suffixes like "-beta" or "-rc1"
        var match = MyRegex().Match(tagName);
        return match.Success && Version.TryParse(match.Value, out var version) ? version : null;
    }

    public void Dispose()
    {
        _httpClient.Dispose();
        GC.SuppressFinalize(this);
    }

    [GeneratedRegex(@"\d+\.\d+\.\d+(\.\d+)?")]
    private static partial Regex MyRegex();
}
