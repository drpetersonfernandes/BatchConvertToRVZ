using System.Globalization;
using System.IO;

namespace BatchConvertToRVZ.services;

/// <summary>
/// Service responsible for file operations and extension management.
/// </summary>
public class FileService
{
    private readonly Action<string> _logMessage;

    // Supported input extensions for conversion (ISO, GCM, WBFS, GCZ, WIA, NKIT.ISO, RAR, 7Z, ZIP)
    private static readonly string[] AllSupportedInputExtensions = [".iso", ".gcm", ".wbfs", ".gcz", ".wia", ".nkit.iso", ".zip", ".7z", ".rar"];

    // Archive extensions
    private static readonly string[] ArchiveExtensions = [".zip", ".7z", ".rar"];

    // Extensions for files inside archives that we want to extract and convert
    private static readonly string[] PrimaryTargetExtensionsInsideArchive = [".iso", ".gcm", ".wbfs", ".rvz", ".gcz", ".wia", ".nkit.iso"];

    // RVZ extensions for verification
    private static readonly string[] RvzExtension = [".rvz"];

    // Extraction input extensions (RVZ, 7Z, RAR, ZIP)
    private static readonly string[] ExtractionInputExtensions = [".rvz", ".zip", ".7z", ".rar"];

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
    /// Gets extraction input file extensions (RVZ, 7Z, RAR, ZIP).
    /// </summary>
    /// <returns>Array of extraction input extensions.</returns>
    public string[] GetExtractionInputExtensions()
    {
        return ExtractionInputExtensions;
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
        var extension = Path.GetExtension(filePath);
        return ArchiveExtensions.Any(ext => ext.Equals(extension, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Determines whether the specified file is a supported input file for conversion.
    /// Handles compound extensions like .nkit.iso correctly.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>true if the file is supported; otherwise, false.</returns>
    public bool IsSupportedInputFile(string filePath)
    {
        var fileName = Path.GetFileName(filePath);

        // Check for compound extensions first (e.g., .nkit.iso, .nkit.gcz) - must be checked before simple extensions
        if (fileName.EndsWith(".nkit.iso", StringComparison.OrdinalIgnoreCase) ||
            fileName.EndsWith(".nkit.gcz", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        // Check for simple extensions
        var extension = Path.GetExtension(filePath);
        return AllSupportedInputExtensions.Any(ext => ext.Equals(extension, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Determines whether the specified file is a supported extraction input file (RVZ, 7Z, RAR, ZIP).
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>true if the file is a supported extraction input; otherwise, false.</returns>
    public bool IsSupportedExtractionInputFile(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        return ExtractionInputExtensions.Any(ext => ext.Equals(extension, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Determines whether the specified file is an RVZ file.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>true if the file is an RVZ; otherwise, false.</returns>
    public bool IsRvzFile(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        return RvzExtension.Any(ext => ext.Equals(extension, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Gets files from a folder that match the specified extensions.
    /// </summary>
    /// <param name="folderPath">The folder path to search.</param>
    /// <param name="extensions">The file extensions to filter by.</param>
    /// <param name="recursive">Whether to search recursively in subdirectories.</param>
    /// <returns>Array of file paths matching the extensions.</returns>
    public string[] GetFilesFromFolder(string folderPath, string[] extensions, bool recursive = false)
    {
        try
        {
            var searchOption = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(folderPath, "*.*", searchOption)
                .Where(file =>
                {
                    var fileName = Path.GetFileName(file);
                    var extension = Path.GetExtension(file);

                    // Check for compound extensions first (e.g., .nkit.iso)
                    if (fileName.EndsWith(".nkit.iso", StringComparison.OrdinalIgnoreCase))
                    {
                        return extensions.Any(static ext => ext.Equals(".nkit.iso", StringComparison.OrdinalIgnoreCase));
                    }

                    return extensions.Any(ext => ext.Equals(extension, StringComparison.OrdinalIgnoreCase));
                })
                .ToArray();

            return files;
        }
        catch (Exception ex)
        {
            _logMessage($"Error reading folder {folderPath}: {ex.Message}");
            return Array.Empty<string>();
        }
    }

    /// <summary>
    /// Attempts to delete a file asynchronously.
    /// </summary>
    /// <param name="filePath">The file path to delete.</param>
    /// <param name="description">Description of the file for logging purposes.</param>
    /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
    /// <returns>true if the file was successfully deleted; otherwise, false.</returns>
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

    /// <summary>
    /// Attempts to delete a directory asynchronously.
    /// </summary>
    /// <param name="dirPath">The directory path to delete.</param>
    /// <param name="description">Description of the directory for logging purposes.</param>
    /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
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

    /// <summary>
    /// Moves a file to a subfolder asynchronously.
    /// </summary>
    /// <param name="sourceFilePath">The source file path to move.</param>
    /// <param name="baseFolder">The base folder containing the subfolder.</param>
    /// <param name="subfolderName">The name of the subfolder to move the file to.</param>
    /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
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

    /// <summary>
    /// Gets the size of a file in bytes.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>The file size in bytes, or 0 if the file doesn't exist or an error occurs.</returns>
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

    /// <summary>
    /// Determines whether the specified file exists.
    /// </summary>
    /// <param name="filePath">The file path to check.</param>
    /// <returns>true if the file exists; otherwise, false.</returns>
    public bool FileExists(string filePath)
    {
        return File.Exists(filePath);
    }

    /// <summary>
    /// Determines whether the specified directory exists.
    /// </summary>
    /// <param name="directoryPath">The directory path to check.</param>
    /// <returns>true if the directory exists; otherwise, false.</returns>
    public bool DirectoryExists(string directoryPath)
    {
        return Directory.Exists(directoryPath);
    }

    /// <summary>
    /// Gets the file name from the specified path.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>The file name including the extension.</returns>
    public string GetFileName(string filePath)
    {
        return Path.GetFileName(filePath);
    }

    /// <summary>
    /// Gets the file name without the extension from the specified path.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>The file name without the extension.</returns>
    public string GetFileNameWithoutExtension(string filePath)
    {
        return Path.GetFileNameWithoutExtension(filePath);
    }

    /// <summary>
    /// Gets the extension of the specified file path.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>The file extension including the leading dot.</returns>
    public string GetExtension(string filePath)
    {
        return Path.GetExtension(filePath);
    }

    /// <summary>
    /// Gets the directory name from the specified path.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>The directory name, or empty string if the path doesn't contain directory information.</returns>
    public string GetDirectoryName(string filePath)
    {
        return Path.GetDirectoryName(filePath) ?? string.Empty;
    }

    /// <summary>
    /// Combines multiple path components into a single path.
    /// </summary>
    /// <param name="paths">An array of path components to combine.</param>
    /// <returns>The combined path.</returns>
    public string CombinePaths(params string[] paths)
    {
        return Path.Combine(paths);
    }

    /// <summary>
    /// Changes the extension of the specified file path.
    /// </summary>
    /// <param name="filePath">The file path to change the extension for.</param>
    /// <param name="newExtension">The new file extension.</param>
    /// <returns>The file path with the new extension.</returns>
    public string ChangeExtension(string filePath, string newExtension)
    {
        return Path.ChangeExtension(filePath, newExtension);
    }

    /// <summary>
    /// Gets the path of the temporary folder for the current system.
    /// </summary>
    /// <returns>The path to the temporary folder.</returns>
    public string GetTempPath()
    {
        return Path.GetTempPath();
    }

    /// <summary>
    /// Gets a random file name.
    /// </summary>
    /// <returns>A random file name.</returns>
    public string GetRandomFileName()
    {
        return Path.GetRandomFileName();
    }

    /// <summary>
    /// Creates a temporary directory with an optional prefix.
    /// </summary>
    /// <param name="prefix">The prefix for the temporary directory name. Defaults to "BatchConvertToRVZ_Temp_".</param>
    /// <returns>The path to the created temporary directory.</returns>
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