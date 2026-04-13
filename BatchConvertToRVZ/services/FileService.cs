using System.Globalization;
using System.IO;

namespace BatchConvertToRVZ.services;

/// <summary>
/// Service responsible for file operations and extension management.
/// </summary>
public class FileService
{
    private readonly Action<string> _logMessage;

    // Supported input extensions
    private static readonly string[] AllSupportedInputExtensions = [".iso", ".gcm", ".wbfs", ".rvz", ".zip", ".7z", ".rar"];

    private static readonly string[] ArchiveExtensions = [".zip", ".7z", ".rar"];
    private static readonly string[] PrimaryTargetExtensionsInsideArchive = [".iso", ".gcm", ".wbfs", ".rvz", ".nkit.iso"];
    private static readonly string[] RvzExtension = [".rvz"];

    public FileService(Action<string> logMessage)
    {
        _logMessage = logMessage;
    }

    /// <summary>
    /// Gets all supported input file extensions.
    /// </summary>
    /// <returns>Array of supported extensions.</returns>
    public string[] GetAllSupportedInputExtensions()
    {
        return AllSupportedInputExtensions;
    }

    /// <summary>
    /// Gets archive file extensions.
    /// </summary>
    /// <returns>Array of archive extensions.</returns>
    public string[] GetArchiveExtensions()
    {
        return ArchiveExtensions;
    }

    /// <summary>
    /// Gets primary target extensions for files inside archives.
    /// </summary>
    /// <returns>Array of target extensions.</returns>
    public string[] GetPrimaryTargetExtensionsInsideArchive()
    {
        return PrimaryTargetExtensionsInsideArchive;
    }

    /// <summary>
    /// Gets RVZ file extensions.
    /// </summary>
    /// <returns>Array of RVZ extensions.</returns>
    public string[] GetRvzExtensions()
    {
        return RvzExtension;
    }

    /// <summary>
    /// Gets a display string of primary target extensions.
    /// </summary>
    /// <returns>Comma-separated list of extensions.</returns>
    public string GetPrimaryTargetExtensionsDisplay()
    {
        return string.Join(", ", PrimaryTargetExtensionsInsideArchive);
    }

    /// <summary>
    /// Determines whether the specified file is an archive.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>true if the file is an archive; otherwise, false.</returns>
    public bool IsArchiveFile(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return ArchiveExtensions.Contains(extension);
    }

    /// <summary>
    /// Determines whether the specified file is a supported input file.
    /// Handles compound extensions like .nkit.iso correctly.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>true if the file is supported; otherwise, false.</returns>
    public bool IsSupportedInputFile(string filePath)
    {
        var fileName = Path.GetFileName(filePath).ToLowerInvariant();

        // Check for compound extensions first (e.g., .nkit.iso)
        foreach (var ext in AllSupportedInputExtensions)
        {
            if (fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        // Fallback to standard extension check
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return AllSupportedInputExtensions.Contains(extension);
    }

    /// <summary>
    /// Determines whether the specified file is an RVZ file.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>true if the file is an RVZ; otherwise, false.</returns>
    public bool IsRvzFile(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return RvzExtension.Contains(extension);
    }

    public string[] GetFilesFromFolder(string folderPath, string[] extensions, bool recursive = false)
    {
        try
        {
            var searchOption = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(folderPath, "*.*", searchOption)
                .Where(file => extensions.Contains(Path.GetExtension(file).ToLowerInvariant()))
                .ToArray();

            return files;
        }
        catch (Exception ex)
        {
            _logMessage($"Error reading folder {folderPath}: {ex.Message}");
            return Array.Empty<string>();
        }
    }

    public async Task<bool> TryDeleteFileAsync(string filePath, string description, CancellationToken cancellationToken)
    {
        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (File.Exists(filePath))
            {
                // File.Delete is synchronous and doesn't support cancellation directly
                // We can only check cancellation before and after the operation
                await Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    File.Delete(filePath);
                    cancellationToken.ThrowIfCancellationRequested();
                }, cancellationToken);

                _logMessage($"Deleted {description}: {Path.GetFileName(filePath)}");
                return true;
            }

            return false;
        }
        catch (OperationCanceledException)
        {
            _logMessage($"Delete operation cancelled for {description}: {Path.GetFileName(filePath)}");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Failed to delete {description} {Path.GetFileName(filePath)}: {ex.Message}");
            return false;
        }
    }

    public async Task TryDeleteDirectoryAsync(string dirPath, string description, CancellationToken cancellationToken = default)
    {
        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (Directory.Exists(dirPath))
            {
                await Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Directory.Delete(dirPath, true);
                    cancellationToken.ThrowIfCancellationRequested();
                }, cancellationToken);

                _logMessage($"Deleted {description}");
            }
        }
        catch (OperationCanceledException)
        {
            _logMessage($"Delete directory operation cancelled for {description}");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Failed to delete {description}: {ex.Message}");
        }
    }

    public async Task MoveFileToSubfolderAsync(string sourceFilePath, string baseFolder, string subfolderName, CancellationToken cancellationToken = default)
    {
        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            var fileName = Path.GetFileName(sourceFilePath);
            var subfolderPath = Path.Combine(baseFolder, subfolderName);

            if (!Directory.Exists(subfolderPath))
            {
                Directory.CreateDirectory(subfolderPath);
            }

            var destinationPath = Path.Combine(subfolderPath, fileName);

            if (File.Exists(destinationPath))
            {
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
                var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                var extension = Path.GetExtension(fileName);
                destinationPath = Path.Combine(subfolderPath, $"{nameWithoutExt}_{timestamp}{extension}");
            }

            await Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                File.Move(sourceFilePath, destinationPath);
                cancellationToken.ThrowIfCancellationRequested();
            }, cancellationToken);

            _logMessage($"Moved {fileName} to {subfolderName} folder.");
        }
        catch (OperationCanceledException)
        {
            _logMessage($"Move file operation cancelled for {Path.GetFileName(sourceFilePath)}");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Failed to move file to {subfolderName} folder: {ex.Message}");
        }
    }

    public long GetFileSize(string filePath)
    {
        try
        {
            return new FileInfo(filePath).Length;
        }
        catch
        {
            return 0;
        }
    }

    public bool FileExists(string filePath)
    {
        return File.Exists(filePath);
    }

    public bool DirectoryExists(string directoryPath)
    {
        return Directory.Exists(directoryPath);
    }

    public string GetFileName(string filePath)
    {
        return Path.GetFileName(filePath);
    }

    public string GetFileNameWithoutExtension(string filePath)
    {
        return Path.GetFileNameWithoutExtension(filePath);
    }

    public string GetExtension(string filePath)
    {
        return Path.GetExtension(filePath);
    }

    public string GetDirectoryName(string filePath)
    {
        return Path.GetDirectoryName(filePath) ?? string.Empty;
    }

    public string CombinePaths(params string[] paths)
    {
        return Path.Combine(paths);
    }

    public string ChangeExtension(string filePath, string newExtension)
    {
        return Path.ChangeExtension(filePath, newExtension);
    }

    public string GetTempPath()
    {
        return Path.GetTempPath();
    }

    public string GetRandomFileName()
    {
        return Path.GetRandomFileName();
    }

    public string CreateTempDirectory(string prefix = "BatchConvertToRVZ_Temp_")
    {
        var tempDir = Path.Combine(GetTempPath(), prefix + GetRandomFileName());
        Directory.CreateDirectory(tempDir);
        return tempDir;
    }

    /// <summary>
    /// Gets the base file name without game image extensions.
    /// Handles compound extensions like .nkit.iso correctly.
    /// </summary>
    /// <summary>
    /// Gets the base file name without game image extensions.
    /// Handles compound extensions like .nkit.iso correctly.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>The base file name without game extensions.</returns>
    public string GetBaseFileNameWithoutGameExtension(string filePath)
    {
        var fileName = Path.GetFileName(filePath);

        // Handle compound extension .nkit.iso explicitly first as it's the only compound one
        if (fileName.EndsWith(".nkit.iso", StringComparison.OrdinalIgnoreCase))
        {
            return fileName[..^9]; // Remove .nkit.iso
        }

        // Handle other game image extensions from our supported list
        foreach (var ext in PrimaryTargetExtensionsInsideArchive)
        {
            if (fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
            {
                return fileName[..^ext.Length];
            }
        }

        // Fallback to standard behavior if no target extension matched
        return Path.GetFileNameWithoutExtension(fileName);
    }
}