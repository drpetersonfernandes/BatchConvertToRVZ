using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace BatchConvertToRVZ.Models;

/// <summary>
/// Represents a file item in the conversion or verification list.
/// Implements INotifyPropertyChanged for data binding support.
/// </summary>
public class FileItem : INotifyPropertyChanged
{
    private bool _isSelected = true;
    private string _fileName = string.Empty;
    private string _fullPath = string.Empty;
    private long _fileSize;
    private string _displaySize = string.Empty;

    /// <summary>
    /// Gets or sets a value indicating whether the file is selected for processing.
    /// </summary>
    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value) return;

            _isSelected = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    /// Gets or sets the name of the file.
    /// </summary>
    public string FileName
    {
        get => _fileName;
        set
        {
            if (_fileName == value) return;

            _fileName = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    /// Gets or sets the full path to the file.
    /// </summary>
    public string FullPath
    {
        get => _fullPath;
        set
        {
            if (_fullPath == value) return;

            _fullPath = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    /// Gets or sets the size of the file in bytes.
    /// When set, automatically updates the <see cref="DisplaySize"/> property.
    /// </summary>
    public long FileSize
    {
        get => _fileSize;
        set
        {
            if (_fileSize == value) return;

            _fileSize = value;
            OnPropertyChanged();
            DisplaySize = FormatSize(value);
        }
    }

    /// <summary>
    /// Gets or sets the human-readable file size string (e.g., "1.5 GB").
    /// </summary>
    public string DisplaySize
    {
        get => _displaySize;
        set
        {
            if (_displaySize == value) return;

            _displaySize = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    /// Occurs when a property value changes.
    /// </summary>
    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>
    /// Raises the <see cref="PropertyChanged"/> event.
    /// </summary>
    /// <param name="propertyName">Name of the property that changed.</param>
    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    /// <summary>
    /// Formats a byte size into a human-readable string.
    /// </summary>
    /// <param name="bytes">The size in bytes.</param>
    /// <returns>A formatted string (e.g., "1.5 GB").</returns>
    private static string FormatSize(long bytes)
    {
        string[] suffix = ["B", "KB", "MB", "GB", "TB"];
        int i;
        double dblSByte = bytes;
        for (i = 0; i < suffix.Length - 1 && bytes >= 1024; i++, bytes /= 1024)
        {
            dblSByte = bytes / 1024.0;
        }

        return $"{dblSByte:0.##} {suffix[i]}";
    }
}