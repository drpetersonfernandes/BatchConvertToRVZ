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
    private static readonly SocketsHttpHandler SharedHandler = new()
    {
        PooledConnectionLifetime = TimeSpan.FromMinutes(2)
    };

    private static int _instanceCount;

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

        Interlocked.Increment(ref _instanceCount);
        _httpClient = new HttpClient(SharedHandler, false);

        // Use Bearer token for authentication as required by the Stats API
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
    }

    /// <summary>
    /// Sends usage statistics to the API.
    /// </summary>
    /// <returns>A task representing the asynchronous operation.</returns>
    public async Task<bool> SendUsageStatsAsync()
    {
        try
        {
            var payload = new
            {
                applicationId = _applicationId,
                version = _applicationVersion
            };

            var response = await _httpClient.PostAsJsonAsync(_apiUrl, payload);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            // Silently fail to not interrupt the user experience
            return false;
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

        if (Interlocked.Decrement(ref _instanceCount) == 0)
        {
            SharedHandler.Dispose();
        }

        GC.SuppressFinalize(this);
    }
}
