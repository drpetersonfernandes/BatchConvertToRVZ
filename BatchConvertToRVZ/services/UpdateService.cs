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
    // Shared static HttpClient handler to prevent socket exhaustion.
    // Properly disposed when the application exits.
    private static readonly Lazy<SocketsHttpHandler> SharedHandler = new(static () => new SocketsHttpHandler
    {
        PooledConnectionLifetime = TimeSpan.FromMinutes(2)
    });

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
        _httpClient = new HttpClient(SharedHandler.Value, false);

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

            var currentVersion = NormalizeVersion(GetCurrentVersion());
            var latestVersion = NormalizeVersion(ParseVersionFromTag(latestRelease.TagName));

            if (latestVersion != null && latestVersion > currentVersion)
            {
                return (true, latestRelease);
            }
        }
        catch (Exception ex)
        {
            // Unexpected errors should be reported for debugging
            // Use the global exception handler through App domain
            System.Diagnostics.Debug.WriteLine($"Unexpected error in UpdateService: {ex}");
            throw; // Re-throw to let global handler catch and report
        }

        return (false, null);
    }

    /// <summary>
    /// Normalizes a version to always have 4 components (Major, Minor, Build, Revision).
    /// This ensures consistent comparisons between tags like "1.8.1" and assembly versions like "1.8.1.0".
    /// </summary>
    private static Version? NormalizeVersion(Version? v)
    {
        if (v == null) return null;

        return new Version(
            v.Major,
            v.Minor,
            v.Build != -1 ? v.Build : 0,
            v.Revision != -1 ? v.Revision : 0
        );
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
    /// <param name="tagName">The tag name (e.g., "v1.2.3", "release-1.2.3", "v1.2", "v1.7.1-beta.1").</param>
    /// <returns>A <see cref="Version"/> object or null if parsing fails.</returns>
    private static Version? ParseVersionFromTag(string tagName)
    {
        if (string.IsNullOrWhiteSpace(tagName))
            return null;

        // Trim whitespace and leading 'v'/'V' prefix commonly used in tags
        var tag = tagName.Trim();
        if (tag.Length > 0 && tag[0] is 'v' or 'V')
        {
            tag = tag[1..];
            // Check if tag became empty after removing prefix
            if (string.IsNullOrEmpty(tag))
                return null;
        }

        // Extract numeric version: supports 2, 3, or 4 segment versions
        // e.g., "1.2", "1.2.3", "1.2.3.4"
        var match = VersionCoreRegex().Match(tag);
        if (!match.Success || !Version.TryParse(match.Value, out var version))
            return null;

        // Strip pre-release suffixes (e.g., "-beta", "-rc1", "-alpha.2")
        // by parsing only the numeric core. This avoids Version.TryParse
        // misinterpreting "-beta" as part of the build number.
        return version;
    }

    [GeneratedRegex(@"\d+\.\d+(\.\d+)?(\.\d+)?")]
    private static partial Regex VersionCoreRegex();

    public void Dispose()
    {
        _httpClient.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Disposes the shared handler when the application exits.
    /// Call this method during application shutdown.
    /// </summary>
    public static void DisposeSharedHandler()
    {
        if (SharedHandler.IsValueCreated)
        {
            SharedHandler.Value.Dispose();
        }
    }
}
