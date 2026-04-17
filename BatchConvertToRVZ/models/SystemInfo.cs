namespace BatchConvertToRVZ.Models;

/// <summary>
/// Model representing system information for bug reports
/// </summary>
public class SystemInfo
{
    // === Environment Details ===

    /// <summary>
    /// Gets or sets the date and time when the bug report was generated.
    /// </summary>
    public string Date { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the name of the application.
    /// </summary>
    public string ApplicationName { get; set; } = string.Empty;

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

    /// <summary>
    /// Gets or sets the number of processors.
    /// </summary>
    public int ProcessorCount { get; set; }

    /// <summary>
    /// Gets or sets the base directory of the application.
    /// </summary>
    public string BaseDirectory { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the system's temporary path.
    /// </summary>
    public string TempPath { get; set; } = string.Empty;

    // === Error Details ===

    /// <summary>
    /// Gets or sets the error message or bug report description.
    /// </summary>
    public string Message { get; set; } = string.Empty;

    // === Exception Details ===

    /// <summary>
    /// Gets or sets the type of the exception that occurred.
    /// </summary>
    public string ExceptionType { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the exception message.
    /// </summary>
    public string ExceptionMessage { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the exception source.
    /// </summary>
    public string ExceptionSource { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the exception stack trace.
    /// </summary>
    public string StackTrace { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the detailed exception information including full exception chain.
    /// </summary>
    public string ExceptionDetails { get; set; } = string.Empty;
}