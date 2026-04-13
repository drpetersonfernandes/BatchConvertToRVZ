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
    // Properly disposed when the application exits.
    private static readonly Lazy<SocketsHttpHandler> SharedHandler = new(static () => new SocketsHttpHandler
    {
        PooledConnectionLifetime = TimeSpan.FromMinutes(2)
    });

    private readonly HttpClient _httpClient;
    private readonly string _apiUrl;
    private readonly string _apiKey;
    private readonly string _applicationId;
    private readonly string _applicationVersion;

    /// <summary>
    /// Initializes a new instance of the <see cref="StatsService"/> class.
    /// </summary>
    /// <param name="apiUrl">The URL of the Stats API.</param>
    /// <param name="apiKey">The API key for authentication.</param>
    /// <param name="applicationId">The unique identifier for the application.</param>
    public StatsService(string apiUrl, string apiKey, string applicationId)
    {
        _apiUrl = apiUrl;
        _apiKey = apiKey;
        _applicationId = applicationId;
        _applicationVersion = GetApplicationVersion();

        _httpClient = new HttpClient(SharedHandler.Value, false);

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
            var statusCode = (int)response.StatusCode;
            throw statusCode switch
            {
                429 => new HttpRequestException($"Stats API Rate Limit: This IP has already reported stats for '{_applicationId}' within the rate limit period (usually 1 hour)."),
                400 => new HttpRequestException($"Stats API Bad Request (400): The request was malformed or missing required fields. Response: {content}"),
                401 => new HttpRequestException("Stats API Unauthorized (401): Invalid or missing API key."),
                403 => new HttpRequestException("Stats API Forbidden (403): API key does not have permission to access this resource."),
                404 => new HttpRequestException($"Stats API Not Found (404): The requested endpoint '{_apiUrl}' does not exist."),
                500 => new HttpRequestException($"Stats API Server Error (500): The server encountered an internal error. Response: {content}"),
                502 => new HttpRequestException("Stats API Bad Gateway (502): The server received an invalid response from an upstream server."),
                503 => new HttpRequestException($"Stats API Service Unavailable (503): The server is temporarily unavailable. Response: {content}"),
                _ => new HttpRequestException($"Stats API failed with status {response.StatusCode}: {content}")
            };
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
