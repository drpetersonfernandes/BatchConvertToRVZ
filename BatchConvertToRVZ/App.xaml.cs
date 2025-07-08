﻿using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using SevenZip;

namespace BatchConvertToRVZ;

public partial class App : IDisposable
{
    // Bug Report API configuration
    private const string BugReportApiUrl = "https://www.purelogiccode.com/bugreport/api/send-bug-report";
    private const string BugReportApiKey = "hjh7yu6t56tyr540o9u8767676r5674534453235264c75b6t7ggghgg76trf564e";
    private const string ApplicationName = "BatchConvertToRVZ";

    private readonly BugReportService? _bugReportService;

    public App()
    {
        // Initialize the bug report service
        _bugReportService = new BugReportService(BugReportApiUrl, BugReportApiKey, ApplicationName);

        // Set up global exception handling
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        DispatcherUnhandledException += App_DispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

        // Initialize SevenZipSharp library path
        InitializeSevenZipSharp();

        // Register the Exit event handler
        Exit += App_Exit;
    }

    private void App_Exit(object sender, ExitEventArgs e)
    {
        // Dispose of the shared BugReportService instance
        _bugReportService?.Dispose();

        // Unregister event handlers to prevent memory leaks
        AppDomain.CurrentDomain.UnhandledException -= CurrentDomain_UnhandledException;
        DispatcherUnhandledException -= App_DispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException -= TaskScheduler_UnobservedTaskException;
    }

    private void InitializeSevenZipSharp()
    {
        try
        {
            // Determine the path to the 7z dll based on the process architecture.
            var dllName = Environment.Is64BitProcess ? "7z_x64.dll" : "7z_x86.dll";
            var dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, dllName);

            if (File.Exists(dllPath))
            {
                SevenZipBase.SetLibraryPath(dllPath);
            }
            else
            {
                // Notify developer
                // If the specific DLL is not found, log an error. Extraction will likely fail.
                var errorMessage = $"Could not find the required 7-Zip library: {dllName} in {AppDomain.CurrentDomain.BaseDirectory}";

                if (_bugReportService != null)
                {
                    _ = _bugReportService.SendBugReportAsync(errorMessage);
                }
            }
        }
        catch (Exception ex)
        {
            // Notify developer
            if (_bugReportService != null)
            {
                _ = _bugReportService.SendBugReportAsync(ex.Message);
            }
        }
    }

    private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
        {
            ReportException(exception, "AppDomain.UnhandledException");
        }
    }

    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        ReportException(e.Exception, "Application.DispatcherUnhandledException");
        e.Handled = true;
    }

    private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        ReportException(e.Exception, "TaskScheduler.UnobservedTaskException");
        e.SetObserved();
    }

    private async void ReportException(Exception exception, string source)
    {
        try
        {
            var message = BuildExceptionReport(exception, source);

            // Notify developer
            if (_bugReportService != null)
            {
                await _bugReportService.SendBugReportAsync(message);
            }
        }
        catch
        {
            // Silently ignore any errors in the reporting process
        }
    }

    private string BuildExceptionReport(Exception exception, string source)
    {
        var sb = new StringBuilder();
        sb.AppendLine(CultureInfo.InvariantCulture, $"Error Source: {source}");
        sb.AppendLine(CultureInfo.InvariantCulture, $"Date and Time: {DateTime.Now}");
        sb.AppendLine(CultureInfo.InvariantCulture, $"OS Version: {Environment.OSVersion}");
        sb.AppendLine(CultureInfo.InvariantCulture, $".NET Version: {Environment.Version}");
        sb.AppendLine();

        // Add exception details
        sb.AppendLine("Exception Details:");
        AppendExceptionDetails(sb, exception);

        return sb.ToString();
    }

    private void AppendExceptionDetails(StringBuilder sb, Exception exception, int level = 0)
    {
        while (true)
        {
            var indent = new string(' ', level * 2);

            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Type: {exception.GetType().FullName}");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Message: {exception.Message}");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Source: {exception.Source}");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}StackTrace:");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}{exception.StackTrace}");

            // If there's an inner exception, include it too
            if (exception.InnerException != null)
            {
                sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Inner Exception:");
                exception = exception.InnerException;
                level += 1;
                continue;
            }

            break;
        }
    }

    /// <inheritdoc />
    /// <summary>
    /// Releases all resources used by the App.
    /// </summary>
    public void Dispose()
    {
        // Dispose the BugReportService if it exists
        _bugReportService?.Dispose();

        // Unregister event handlers to prevent memory leaks
        AppDomain.CurrentDomain.UnhandledException -= CurrentDomain_UnhandledException;
        DispatcherUnhandledException -= App_DispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException -= TaskScheduler_UnobservedTaskException;

        // Suppress finalization
        GC.SuppressFinalize(this);
    }
}
