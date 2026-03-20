using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Reflection;

namespace BatchConvertToRVZ.services;

/// <summary>
/// Service responsible for sending application usage statistics to the Stats API.
/// </summary>
public class StatsService : IDisposable
{
    // Shared static HttpClient handler to prevent socket exhaustion.
    // Not disposed explicitly — lives for the app lifetime and is cleaned up by the finalizer on exit.
    private static readonly SocketsHttpHandler SharedHandler = new()
    {
        PooledConnectionLifetime = TimeSpan.FromMinutes(2)
    };

    private readonly HttpClient _httpClient;
    private readonly string _apiUrl;
    private readonly string _apiKey;
    private readonly string _applicationId;
    private readonly string _applicationVersion;

    public StatsService(string apiUrl, string apiKey, string applicationId)
    {
        _apiUrl = apiUrl;
        _apiKey = apiKey;
        _applicationId = applicationId;
        _applicationVersion = GetApplicationVersion();

        _httpClient = new HttpClient(SharedHandler, false);

        // Use Bearer token for authentication as required by the Stats API
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
    }

    /// <summary>
    /// Sends usage statistics to the API.
    /// </summary>
    /// <returns>A task representing the asynchronous operation.</returns>
    public async Task SendUsageStatsAsync()
    {
        var payload = new
        {
            applicationId = _applicationId,
            version = _applicationVersion
        };

        var response = await _httpClient.PostAsJsonAsync(_apiUrl, payload);

        if (!response.IsSuccessStatusCode)
        {
            var content = await response.Content.ReadAsStringAsync();

            // Provide specific details for common failures
            if ((int)response.StatusCode == 429)
            {
                throw new HttpRequestException($"Stats API Rate Limit: This IP has already reported stats for '{_applicationId}' within the rate limit period (usually 1 hour).");
            }

            throw new HttpRequestException($"Stats API failed with status {response.StatusCode}: {content}");
        }
    }

    private static string GetApplicationVersion()
    {
        var assembly = Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();
        var version = assembly.GetName().Version;
        return version?.ToString() ?? "Unknown";
    }

    public void Dispose()
    {
        _httpClient.Dispose();
        GC.SuppressFinalize(this);
    }
}
