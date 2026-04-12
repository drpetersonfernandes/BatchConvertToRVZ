namespace BatchConvertToRVZ.models;

/// <summary>
/// Model representing system information for bug reports
/// </summary>
public class SystemInfo
{
    /// <summary>
    /// Gets or sets the error message or bug report description.
    /// </summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the name of the application.
    /// </summary>
    public string ApplicationName { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the date and time when the bug report was generated.
    /// </summary>
    public string Date { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the version of the application.
    /// </summary>
    public string ApplicationVersion { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the operating system version.
    /// </summary>
    public string OsVersion { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the processor architecture.
    /// </summary>
    public string Architecture { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the operating system bitness (32-bit or 64-bit).
    /// </summary>
    public string Bitness { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the friendly Windows version name.
    /// </summary>
    public string WindowsVersion { get; set; } = string.Empty;
}