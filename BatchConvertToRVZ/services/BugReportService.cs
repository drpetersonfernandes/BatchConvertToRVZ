using System.Net.Http;
using System.Net.Http.Json;
using System.Reflection;
using System.Runtime.InteropServices;
using BatchConvertToRVZ.models;

namespace BatchConvertToRVZ.services;

/// <inheritdoc />
/// <summary>
/// Service responsible for sending bug reports to the BugReport API
/// </summary>
public class BugReportService : IDisposable
{
    private readonly HttpClient _httpClient = new();
    private readonly string _apiUrl;
    private readonly string _apiKey;
    private readonly string _applicationName;
    private readonly string _applicationVersion;

    public BugReportService(string apiUrl, string apiKey, string applicationName)
    {
        _apiUrl = apiUrl;
        _apiKey = apiKey;
        _applicationName = applicationName;
        _applicationVersion = GetApplicationVersion();

        // Set default headers once in the constructor for thread safety
        _httpClient.DefaultRequestHeaders.Add("X-API-KEY", _apiKey);
    }

    /// <summary>
    /// Sends a bug report to the API with system information
    /// </summary>
    /// <param name="message">The error message or bug report</param>
    /// <returns>A task representing the asynchronous operation</returns>
    public async Task<bool> SendBugReportAsync(string message)
    {
        try
        {
            // Get system information
            var systemInfo = GetSystemInfo(message);

            // Create the request payload using the SystemInfo model for type safety
            var content = JsonContent.Create(systemInfo);

            // Send the request
            var response = await _httpClient.PostAsync(_apiUrl, content);

            // Return true if successful
            return response.IsSuccessStatusCode;
        }
        catch
        {
            // Silently fail if there's an exception
            return false;
        }
    }

    /// <summary>
    /// Gets the current application version from the assembly
    /// </summary>
    private static string GetApplicationVersion()
    {
        var assembly = Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();
        var version = assembly.GetName().Version;
        return version?.ToString() ?? "Unknown";
    }

    /// <summary>
    /// Gets system information for the bug report
    /// </summary>
    /// <param name="message">The error message or bug report</param>
    private SystemInfo GetSystemInfo(string message)
    {
        return new SystemInfo
        {
            Message = message,
            ApplicationName = _applicationName,
            Date = DateTime.Now.ToString("yyyy/M/d tt h:mm:ss", System.Globalization.CultureInfo.InvariantCulture),
            ApplicationVersion = _applicationVersion,
            OsVersion = Environment.OSVersion.ToString(),
            Architecture = RuntimeInformation.ProcessArchitecture.ToString(),
            Bitness = Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit",
            WindowsVersion = GetWindowsVersion()
        };
    }

    /// <summary>
    /// Gets a friendly Windows version name
    /// </summary>
    private static string GetWindowsVersion()
    {
        var osVersion = Environment.OSVersion;
        if (osVersion.Platform == PlatformID.Win32NT)
        {
            var version = osVersion.Version;
            switch (version.Major)
            {
                case 10 when version.Build >= 22000:
                    return "Windows 11";
                case 10:
                    return "Windows 10";
                case 6 when version.Minor == 3:
                    return "Windows 8.1";
                case 6 when version.Minor == 2:
                    return "Windows 8";
                case 6 when version.Minor == 1:
                    return "Windows 7";
            }
        }

        return "Unknown Windows Version";
    }

    public void Dispose()
    {
        // Dispose the HttpClient to release resources
        _httpClient.Dispose();

        // Suppress finalization
        GC.SuppressFinalize(this);
    }
}
