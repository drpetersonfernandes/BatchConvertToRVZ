using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Text;
using SharpCompress.Archives;

namespace BatchConvertToRVZ.services;

/// <summary>
/// Service responsible for extracting RVZ files to ISO format.
/// Supports direct RVZ files and RVZ files inside archives.
/// </summary>
public class ExtractionService
{
    private readonly Action<string> _logMessage;
    private readonly Func<string, Exception?, Task> _reportBugAsync;
    private readonly FileService _fileService;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExtractionService"/> class.
    /// </summary>
    /// <param name="logMessage">Action to log messages.</param>
    /// <param name="reportBugAsync">Function to report bugs asynchronously with optional exception.</param>
    /// <param name="fileService">FileService instance to use for file operations.</param>
    public ExtractionService(
        Action<string> logMessage,
        Func<string, Exception?, Task> reportBugAsync,
        FileService fileService)
    {
        _logMessage = logMessage;
        _reportBugAsync = reportBugAsync;
        _fileService = fileService;
    }

    // Valid output formats for extraction
    private static readonly HashSet<string> ValidOutputFormats = new(StringComparer.OrdinalIgnoreCase)
    {
        "iso",
        "wbfs",
        "gcz",
        "wia"
    };

    /// <summary>
    /// Performs batch extraction of RVZ files to ISO format.
    /// </summary>
    /// <param name="dolphinToolPath">Path to the DolphinTool executable.</param>
    /// <param name="files">Array of file paths to extract.</param>
    /// <param name="outputFolder">Output folder for extracted files.</param>
    /// <param name="deleteFiles">Whether to delete original files after successful extraction.</param>
    /// <param name="outputFormat">Output format (iso, wbfs, gcz, wia).</param>
    /// <param name="updateProgress">Callback to update progress.</param>
    /// <param name="incrementSuccess">Callback to increment success count.</param>
    /// <param name="incrementFailure">Callback to increment failure count.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public async Task PerformBatchExtractionAsync(
        string dolphinToolPath,
        string[] files,
        string outputFolder,
        bool deleteFiles,
        string outputFormat,
        Action<int, int, string> updateProgress,
        Action<int> incrementSuccess,
        Action<int> incrementFailure,
        CancellationToken cancellationToken)
    {
        // Validate output format
        if (!ValidOutputFormats.Contains(outputFormat))
        {
            _logMessage($"Error: Invalid output format '{outputFormat}'. Valid formats are: {string.Join(", ", ValidOutputFormats)}");
            return;
        }

        try
        {
            _logMessage("Preparing for batch extraction...");

            var totalFilesToProcess = files.Length;
            _logMessage($"Processing {totalFilesToProcess} selected files.");

            if (totalFilesToProcess == 0)
            {
                _logMessage("No files selected for extraction.");
                return;
            }

            var filesProcessedCount = 0;

            _logMessage("Processing files sequentially.");
            foreach (var inputFile in files)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var fileName = Path.GetFileName(inputFile);
                _logMessage($"Processing: {fileName}");

                var success = await ProcessFileAsync(
                    dolphinToolPath,
                    inputFile,
                    outputFolder,
                    deleteFiles,
                    outputFormat,
                    cancellationToken);

                if (success)
                {
                    incrementSuccess(1);
                    _logMessage($"Extraction successful: {fileName}");
                }
                else
                {
                    incrementFailure(1);
                    _logMessage($"Extraction failed: {fileName}");
                }

                filesProcessedCount++;
                updateProgress(filesProcessedCount, totalFilesToProcess, fileName);
            }
        }
        catch (OperationCanceledException)
        {
            _logMessage("Batch extraction operation was canceled.");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Error during batch extraction: {ex.Message}");
            await _reportBugAsync("Error during batch extraction operation", ex);
        }
    }

    private async Task<bool> ProcessFileAsync(
        string dolphinToolPath,
        string inputFile,
        string outputFolder,
        bool deleteOriginal,
        string outputFormat,
        CancellationToken cancellationToken)
    {
        var fileName = Path.GetFileName(inputFile);
        var inputExtension = Path.GetExtension(inputFile).ToLowerInvariant();

        try
        {
            if (_fileService.GetArchiveExtensions().Contains(inputExtension))
            {
                return await ProcessArchiveFileAsync(
                    dolphinToolPath,
                    inputFile,
                    outputFolder,
                    deleteOriginal,
                    outputFormat,
                    cancellationToken);
            }
            else
            {
                return await ExtractSingleFileAsync(
                    dolphinToolPath,
                    inputFile,
                    outputFolder,
                    deleteOriginal,
                    outputFormat,
                    cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            _logMessage($"Processing canceled: {fileName}");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Error processing {fileName}: {ex.Message}");
            return false;
        }
    }

    private async Task<bool> ProcessArchiveFileAsync(
        string dolphinToolPath,
        string archivePath,
        string outputFolder,
        bool deleteOriginal,
        string outputFormat,
        CancellationToken cancellationToken)
    {
        var archiveFileName = Path.GetFileName(archivePath);

        try
        {
            _logMessage($"Extracting archive: {archiveFileName}");

            var extractionResult = await ExtractRvzFromArchiveAsync(archivePath, cancellationToken);
            if (!extractionResult.Success)
            {
                _logMessage($"Failed to extract {archiveFileName}: {extractionResult.ErrorMessage}");
                return false;
            }

            var extractedFilePath = extractionResult.FilePath;
            var tempDir = extractionResult.TempDir;

            try
            {
                var success = await ExtractSingleFileAsync(
                    dolphinToolPath,
                    extractedFilePath,
                    outputFolder,
                    false, // Don't delete extracted file - we'll clean up temp dir
                    outputFormat,
                    cancellationToken);

                if (success && deleteOriginal)
                {
                    await TryDeleteFile(archivePath, "original archive");
                }

                return success;
            }
            finally
            {
                await TryDeleteDirectory(tempDir, "temporary extraction directory");
            }
        }
        catch (OperationCanceledException)
        {
            _logMessage($"Archive processing canceled: {archiveFileName}");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Error processing archive {archiveFileName}: {ex.Message}");
            return false;
        }
    }

    private async Task<(bool Success, string FilePath, string TempDir, string ErrorMessage)> ExtractRvzFromArchiveAsync(string archivePath, CancellationToken cancellationToken)
    {
        var tempDir = string.Empty;

        try
        {
            // Create temporary directory for extraction
            tempDir = Path.Combine(Path.GetTempPath(), "BatchConvertToRVZ_Extract_" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);

            _logMessage($"Extracting archive to temporary directory: {tempDir}");

            using var archive = ArchiveFactory.OpenArchive(archivePath);
            var rvzExtensions = _fileService.GetRvzExtensions();

            var entry = archive.Entries.FirstOrDefault(e =>
                e is { IsDirectory: false, Key: not null } &&
                rvzExtensions.Any(ext => e.Key.EndsWith(ext, StringComparison.OrdinalIgnoreCase)));

            if (entry == null)
            {
                var archiveName = Path.GetFileName(archivePath);
                return (false, string.Empty, string.Empty, $"No RVZ file found inside {archiveName}.");
            }

            // Extract the file name from the entry key, handling potential directory separators
            var entryName = entry.Key?.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
            entryName = Path.GetFileName(entryName);

            // If entry name is still invalid, create one from archive name preserving the extension
            if (string.IsNullOrWhiteSpace(entryName))
            {
                var archiveName = Path.GetFileNameWithoutExtension(archivePath);
                var entryExtension = rvzExtensions.FirstOrDefault(ext =>
                    entry.Key?.EndsWith(ext, StringComparison.OrdinalIgnoreCase) == true) ?? ".rvz";
                entryName = $"{archiveName}{entryExtension}";
            }

            var extractedFilePath = Path.Combine(tempDir, entryName);

            await using (var source = await entry.OpenEntryStreamAsync(cancellationToken).ConfigureAwait(false))
            await using (var destination = File.Create(extractedFilePath))
            {
                await source.CopyToAsync(destination, cancellationToken).ConfigureAwait(false);
            }

            _logMessage($"Extracted {entryName} from archive.");
            return (true, extractedFilePath, tempDir, string.Empty);
        }
        catch (Exception ex)
        {
            _logMessage($"Error extracting archive {Path.GetFileName(archivePath)}: {ex.Message}");

            // Clean up on failure
            if (!string.IsNullOrEmpty(tempDir) && Directory.Exists(tempDir))
            {
                try
                {
                    Directory.Delete(tempDir, true);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }

            return (false, string.Empty, string.Empty, $"Failed to extract archive: {ex.Message}");
        }
    }

    private async Task<bool> ExtractSingleFileAsync(
        string dolphinToolPath,
        string inputFile,
        string outputFolder,
        bool deleteOriginal,
        string outputFormat,
        CancellationToken cancellationToken)
    {
        var fileName = Path.GetFileName(inputFile);
        var fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
        var outputFileName = fileNameWithoutExt + "." + outputFormat.ToLowerInvariant();
        var outputFile = Path.Combine(outputFolder, outputFileName);

        try
        {
            _logMessage($"Converting to {outputFormat.ToUpperInvariant()}: {fileName} -> {outputFileName}");

            // Check if output file already exists and delete it
            if (File.Exists(outputFile))
            {
                try
                {
                    File.Delete(outputFile);
                    _logMessage($"Deleted existing output file: {outputFileName}");
                }
                catch (Exception ex)
                {
                    _logMessage($"Failed to delete existing file {outputFileName}: {ex.Message}");
                    return false;
                }
            }

            var success = await ConvertToFormatAsync(
                dolphinToolPath,
                inputFile,
                outputFile,
                outputFormat,
                cancellationToken);

            if (success && deleteOriginal)
            {
                await TryDeleteFile(inputFile, "original file");
            }

            return success;
        }
        catch (OperationCanceledException)
        {
            _logMessage($"Conversion canceled: {fileName}");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Error converting {fileName}: {ex.Message}");
            return false;
        }
    }

    private async Task<bool> ConvertToFormatAsync(
        string dolphinToolPath,
        string inputFile,
        string outputFile,
        string outputFormat,
        CancellationToken cancellationToken)
    {
        using var process = new Process();

        try
        {
            process.StartInfo = new ProcessStartInfo
            {
                FileName = dolphinToolPath,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            process.StartInfo.ArgumentList.Add("convert");
            process.StartInfo.ArgumentList.Add("-i");
            process.StartInfo.ArgumentList.Add(inputFile);
            process.StartInfo.ArgumentList.Add("-o");
            process.StartInfo.ArgumentList.Add(outputFile);
            process.StartInfo.ArgumentList.Add("-f");
            process.StartInfo.ArgumentList.Add(outputFormat.ToLowerInvariant());

            process.EnableRaisingEvents = true;

            var outputQueue = new ConcurrentQueue<string>();
            var errorQueue = new ConcurrentQueue<string>();
            var outputCompleted = new TaskCompletionSource<bool>();
            var errorCompleted = new TaskCompletionSource<bool>();

            process.OutputDataReceived += (_, args) =>
            {
                if (args.Data is null)
                {
                    outputCompleted.TrySetResult(true);
                }
                else
                {
                    outputQueue.Enqueue(args.Data);
                    _logMessage($"[DolphinTool] {args.Data}");
                }
            };

            process.ErrorDataReceived += (_, args) =>
            {
                if (args.Data is null)
                {
                    errorCompleted.TrySetResult(true);
                }
                else
                {
                    errorQueue.Enqueue(args.Data);
                    _logMessage($"[DolphinTool ERROR] {args.Data}");
                }
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            await process.WaitForExitAsync(cancellationToken);
            await Task.WhenAll(outputCompleted.Task, errorCompleted.Task);

            var outputBuilder = new StringBuilder();
            while (outputQueue.TryDequeue(out var line)) outputBuilder.AppendLine(line);
            var output = outputBuilder.ToString();

            var errorBuilder = new StringBuilder();
            while (errorQueue.TryDequeue(out var line)) errorBuilder.AppendLine(line);
            var error = errorBuilder.ToString();

            // Success check: exit code 0 AND file exists
            if (process.ExitCode == 0)
            {
                // Wait for the file to be written (with timeout)
                var fileExists = await WaitForFileExistsAsync(outputFile, cancellationToken);

                if (fileExists)
                {
                    _logMessage($"Successfully converted to {outputFormat.ToUpperInvariant()}: {Path.GetFileName(outputFile)}");
                    return true;
                }
                else
                {
                    _logMessage($"Conversion failed for {Path.GetFileName(inputFile)}. Output file not found after waiting.");
                    if (!string.IsNullOrEmpty(error))
                        _logMessage($"Error output: {error}");
                    return false;
                }
            }
            else
            {
                _logMessage($"Conversion failed for {Path.GetFileName(inputFile)}. Exit code: {process.ExitCode}");
                if (!string.IsNullOrEmpty(output))
                    _logMessage($"Output: {output}");
                if (!string.IsNullOrEmpty(error))
                    _logMessage($"Error output: {error}");
                return false;
            }
        }
        catch (OperationCanceledException)
        {
            if (!process.HasExited) process.Kill(true);
            throw;
        }
    }

    private static async Task<bool> WaitForFileExistsAsync(string filePath, CancellationToken cancellationToken)
    {
        // Wait up to 5 seconds for the file to appear (file system flush delay)
        const int maxAttempts = 50;
        const int delayMs = 100;

        for (var i = 0; i < maxAttempts; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (File.Exists(filePath))
                return true;

            await Task.Delay(delayMs, cancellationToken);
        }

        return false;
    }

    private Task<bool> TryDeleteFile(string filePath, string description)
    {
        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                _logMessage($"Deleted {description}: {Path.GetFileName(filePath)}");
                return Task.FromResult(true);
            }

            return Task.FromResult(false);
        }
        catch (Exception ex)
        {
            _logMessage($"Failed to delete {description} {Path.GetFileName(filePath)}: {ex.Message}");
            return Task.FromResult(false);
        }
    }

    private Task TryDeleteDirectory(string dirPath, string description)
    {
        try
        {
            if (Directory.Exists(dirPath))
            {
                Directory.Delete(dirPath, true);
                _logMessage($"Deleted {description}");
            }
        }
        catch (Exception ex)
        {
            _logMessage($"Failed to delete {description}: {ex.Message}");
        }

        return Task.CompletedTask;
    }
}
