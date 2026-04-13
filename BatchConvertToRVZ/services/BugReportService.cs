using System.Net.Http;
using System.Net.Http.Json;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using BatchConvertToRVZ.Models;

namespace BatchConvertToRVZ.services;

/// <inheritdoc />
/// <summary>
/// Service responsible for sending bug reports to the BugReport API
/// </summary>
public class BugReportService : IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly string _apiUrl;
    private readonly string _apiKey;
    private readonly string _applicationName;
    private readonly string _applicationVersion;

    /// <summary>
    /// Initializes a new instance of the <see cref="BugReportService"/> class.
    /// </summary>
    /// <param name="apiUrl">The URL of the BugReport API.</param>
    /// <param name="apiKey">The API key for authentication.</param>
    /// <param name="applicationName">The name of the application sending bug reports.</param>
    public BugReportService(string apiUrl, string apiKey, string applicationName)
    {
        _apiUrl = apiUrl;
        _apiKey = apiKey;
        _applicationName = applicationName;
        _applicationVersion = GetApplicationVersion();

        // Create a new HttpClient instance that shares the static handler
        // This allows per-instance headers while sharing the connection pool
        _httpClient = new HttpClient(SharedHttpHandler.Instance, false);

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
    /// Sends a bug report to the API with system information including exception details
    /// </summary>
    /// <param name="message">The error message or bug report</param>
    /// <param name="exception">The exception that occurred</param>
    /// <returns>A task representing the asynchronous operation</returns>
    public async Task<bool> SendBugReportAsync(string message, Exception exception)
    {
        try
        {
            // Format the exception details
            var formattedMessage = FormatExceptionMessage(message, exception);

            // Send the bug report
            return await SendBugReportAsync(formattedMessage);
        }
        catch
        {
            // Silently fail if there's an exception
            return false;
        }
    }

    /// <summary>
    /// Formats an exception message for inclusion in a bug report
    /// </summary>
    /// <param name="message">The error message</param>
    /// <param name="exception">The exception that occurred</param>
    /// <returns>A formatted error message with exception details</returns>
    private static string FormatExceptionMessage(string message, Exception exception)
    {
        var sb = new StringBuilder();
        sb.AppendLine(message);
        sb.AppendLine();
        sb.AppendLine("Exception Details:");

        var level = 0;
        var currentException = exception;

        while (currentException != null)
        {
            var indent = new string(' ', level * 2);
            sb.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"{indent}Type: {currentException.GetType().FullName}");
            sb.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"{indent}Message: {currentException.Message}");
            sb.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"{indent}Source: {currentException.Source}");
            sb.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"{indent}StackTrace:");
            sb.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"{indent}{currentException.StackTrace}");

            if (currentException.InnerException != null)
            {
                sb.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"{indent}Inner Exception:");
                currentException = currentException.InnerException;
                level++;
            }
            else
            {
                break;
            }
        }

        return sb.ToString();
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
            ApplicationName = _applicationName,
            Date = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture),
            ApplicationVersion = _applicationVersion,
            OsVersion = Environment.OSVersion.ToString(),
            Architecture = RuntimeInformation.ProcessArchitecture.ToString().ToUpperInvariant(),
            Bitness = Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit",
            WindowsVersion = GetWindowsVersion(),
            Message = message
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
        _httpClient.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Disposes the shared handler when the application exits.
    /// Call this method during application shutdown.
    /// </summary>
    public static void DisposeSharedHandler()
    {
        SharedHttpHandler.Dispose();
    }
}
