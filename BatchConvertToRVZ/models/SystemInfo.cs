namespace BatchConvertToRVZ.models;

/// <summary>
/// Model representing system information for bug reports
/// </summary>
public class SystemInfo
{
    public string Message { get; set; } = string.Empty;
    public string ApplicationName { get; set; } = string.Empty;
    public string Date { get; set; } = string.Empty;
    public string ApplicationVersion { get; set; } = string.Empty;
    public string OsVersion { get; set; } = string.Empty;
    public string Architecture { get; set; } = string.Empty;
    public string Bitness { get; set; } = string.Empty;
    public string WindowsVersion { get; set; } = string.Empty;
}