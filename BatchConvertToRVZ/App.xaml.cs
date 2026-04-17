using System.IO;
using System.Windows;
using System.Windows.Threading;
using BatchConvertToRVZ.services;

namespace BatchConvertToRVZ;

public partial class App
{
    // Bug Report API configuration
    private const string BugReportApiUrl = "https://www.purelogiccode.com/bugreport/api/send-bug-report";
    private const string BugReportApiKey = "hjh7yu6t56tyr540o9u8767676r5674534453235264c75b6t7ggghgg76trf564e";
    private const string ApplicationName = "BatchConvertToRVZ";

    // Stats API configuration
    private const string StatsApiUrl = "https://www.purelogiccode.com/ApplicationStats/stats";
    private const string StatsApiKey = "hjh7yu6t56tyr540o9u8767676r5674534453235264c75b6t7ggghgg76trf564e";
    private const string StatsApplicationId = "BatchConvertToRVZ";

    public static BugReportService? BugReportServiceInstance { get; private set; }
    public static StatsService? StatsServiceInstance { get; private set; }

    public App()
    {
        // Clean up old DLL files from previous versions
        CleanupOldDllFiles();

        // Initialize the services
        BugReportServiceInstance = new BugReportService(BugReportApiUrl, BugReportApiKey, ApplicationName);
        StatsServiceInstance = new StatsService(StatsApiUrl, StatsApiKey, StatsApplicationId);

        // Set up global exception handling
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        DispatcherUnhandledException += App_DispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

        // Register the Exit event handler
        Exit += App_Exit;
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Send usage statistics on application launch
        if (StatsServiceInstance != null)
        {
            _ = Task.Run((Func<Task>)(async () =>
            {
                try
                {
                    await StatsServiceInstance.SendUsageStatsAsync();
                }
                catch (Exception ex)
                {
                    // Silently ignore rate limit exceptions - this is expected behavior
                    // when the user launches the app multiple times within the rate limit window
                    if (ex.Message.Contains("Rate Limit", StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    // Only report other types of exceptions that might indicate actual problems
                    ReportException(ex, "StatsService.OnStartup", false);
                }
            }));
        }
    }

    private static void CleanupOldDllFiles()
    {
        try
        {
            var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var dllFilesToDelete = new[] { "7z_x64.dll", "7z_arm64.dll" };

            foreach (var dllFile in dllFilesToDelete)
            {
                var filePath = Path.Combine(appDirectory, dllFile);
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
        }
        catch
        {
            // Silently ignore any errors during cleanup
        }
    }

    private void App_Exit(object sender, ExitEventArgs e)
    {
        // Dispose of the services
        BugReportServiceInstance?.Dispose();
        StatsServiceInstance?.Dispose();

        // Unregister event handlers to prevent memory leaks
        AppDomain.CurrentDomain.UnhandledException -= CurrentDomain_UnhandledException;
        DispatcherUnhandledException -= App_DispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException -= TaskScheduler_UnobservedTaskException;
    }

    private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
        {
            ReportException(exception, "AppDomain.UnhandledException", true);
        }
    }

    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        var ex = e.Exception;

        // For critical exceptions, let the app crash rather than running in a corrupted state.
        // IO and standard OperationCanceled are usually safe to "handle" if they bubbled up to here.
        if (ex is IOException or TaskCanceledException or OperationCanceledException or UnauthorizedAccessException)
        {
            ReportException(ex, "Application.DispatcherUnhandledException", false);
            MessageBox.Show($"An unexpected but recoverable error occurred: {ex.Message}\n\nThe application will continue to run, but the current operation may have failed.",
                "Recoverable Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            e.Handled = true;
        }
        else
        {
            ReportException(ex, "Application.DispatcherUnhandledException", true);
            // For other unknown/severe exceptions, show a final error message and let it crash to prevent data corruption
            MessageBox.Show($"A fatal error occurred and the application must close: {ex.Message}\n\nA bug report has been sent.",
                "Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);

            // We DON'T set e.Handled = true here, allowing the app to terminate
        }
    }

    private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        ReportException(e.Exception, "TaskScheduler.UnobservedTaskException", false);
        e.SetObserved();
    }

    private void ReportException(Exception exception, string source, bool isFatal)
    {
        try
        {
            var message = $"Error Source: {source}";

            // Notify developer using the exception overload for proper formatting
            if (BugReportServiceInstance != null)
            {
                var reportTask = BugReportServiceInstance.SendBugReportAsync(message, exception);

                if (isFatal)
                {
                    // For fatal exceptions, block the thread to ensure the report is sent before the process terminates.
                    // We use a timeout to prevent hanging forever if the network is down.
                    _ = reportTask.Wait(TimeSpan.FromSeconds(5));
                }
                else
                {
                    // For non-fatal exceptions, fire and forget
                    _ = reportTask;
                }
            }
        }
        catch
        {
            // Silently ignore any errors in the reporting process
        }
    }
}