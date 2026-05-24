using System.IO;

namespace BatchConvertToRVZ.services;

/// <summary>
/// Service responsible for file operations and extension management.
/// </summary>
public class FileService
{
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