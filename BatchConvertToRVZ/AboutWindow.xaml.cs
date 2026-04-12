using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Navigation;

namespace BatchConvertToRVZ;

/// <summary>
/// Interaction logic for the About window.
/// Displays application version and credits information.
/// </summary>
public partial class AboutWindow
{
    /// <summary>
    /// Initializes a new instance of the <see cref="AboutWindow"/> class.
    /// </summary>
    public AboutWindow()
    {
        try
        {
            InitializeComponent();
            AppVersionTextBlock.Text = $"Version: {GetApplicationVersion()}";
        }
        catch (Exception ex)
        {
            // Notify developer if initialization fails
            if (App.BugReportServiceInstance != null)
            {
                _ = App.BugReportServiceInstance.SendBugReportAsync($"Error initializing AboutWindow: {ex.Message}");
            }

            // Notify user and rethrow to prevent window from opening in invalid state
            MessageBox.Show($"Error initializing About window: {ex.Message}",
                "Initialization Error", MessageBoxButton.OK, MessageBoxImage.Error);
            throw;
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        try
        {
            var process = Process.Start(new ProcessStartInfo
            {
                FileName = e.Uri.AbsoluteUri,
                UseShellExecute = true
            });

            process?.Dispose();
        }
        catch (Exception ex)
        {
            // Notify developer
            if (App.BugReportServiceInstance != null)
            {
                _ = App.BugReportServiceInstance.SendBugReportAsync($"Error opening URL: {e.Uri.AbsoluteUri}. Exception: {ex.Message}");
            }

            // Notify user
            Application.Current.Dispatcher.Invoke(() =>
            {
                MessageBox.Show($"Unable to open link: {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            });
        }

        // Mark the event as handled
        e.Handled = true;
    }

    private static string GetApplicationVersion()
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        return version?.ToString() ?? "Unknown";
    }
}