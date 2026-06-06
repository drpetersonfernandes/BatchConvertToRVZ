using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Channels;
using SharpCompress.Archives;

namespace BatchConvertToRVZ.services;

/// <summary>
/// Service responsible for converting game disc images to RVZ format.
/// Supports direct file conversion and extraction from archives.
/// </summary>
public class ConversionService
{
    private readonly Action<string> _logMessage;
    private readonly Func<string, Exception?, Task> _reportBugAsync;
    private readonly FileService _fileService;

    /// <summary>
    /// Initializes a new instance of the <see cref="ConversionService"/> class.
    /// </summary>
    /// <param name="logMessage">Action to log messages.</param>
    /// <param name="reportBugAsync">Function to report bugs asynchronously with optional exception.</param>
    /// <param name="fileService">FileService instance to use for file operations.</param>
    public ConversionService(
        Action<string> logMessage,
        Func<string, Exception?, Task> reportBugAsync,
        FileService fileService)
    {
        _logMessage = logMessage;
        _reportBugAsync = reportBugAsync;
        _fileService = fileService;
    }

    /// <summary>
    /// Performs batch conversion of files to RVZ format.
    /// </summary>
    /// <param name="dolphinToolPath">Path to the DolphinTool executable.</param>
    /// <param name="files">Array of file paths to convert.</param>
    /// <param name="outputFolder">Output folder for converted files.</param>
    /// <param name="deleteFiles">Whether to delete original files after successful conversion.</param>
    /// <param name="compressionMethod">Compression method to use (e.g., "zstd").</param>
    /// <param name="compressionLevel">Compression level (method-specific range).</param>
    /// <param name="blockSize">Block size for compression.</param>
    /// <param name="updateProgress">Callback to update progress.</param>
    /// <param name="incrementSuccess">Callback to increment success count.</param>
    /// <param name="incrementFailure">Callback to increment failure count.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public async Task PerformBatchConversionAsync(
        string dolphinToolPath,
        string[] files,
        string outputFolder,
        bool deleteFiles,
        string compressionMethod,
        int compressionLevel,
        int blockSize,
        Action<int, int, string> updateProgress,
        Action<int> incrementSuccess,
        Action<int> incrementFailure,
        CancellationToken cancellationToken)
    {
        try
        {
            _logMessage("Preparing for batch conversion...");

            var totalFilesToProcess = files.Length;
            _logMessage($"Processing {totalFilesToProcess} selected files.");

            if (totalFilesToProcess == 0)
            {
                _logMessage("No files selected for conversion.");
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
                    compressionMethod,
                    compressionLevel,
                    blockSize,
                    cancellationToken);

                if (success)
                {
                    incrementSuccess(1);
                    _logMessage($"Conversion successful: {fileName}");
                }
                else
                {
                    incrementFailure(1);
                    _logMessage($"Conversion failed: {fileName}");
                }

                filesProcessedCount++;
                updateProgress(filesProcessedCount, totalFilesToProcess, fileName);
            }
        }
        catch (OperationCanceledException)
        {
            _logMessage("Batch conversion operation was canceled.");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Error during batch conversion: {ex.Message}");
            await _reportBugAsync("Error during batch conversion operation", ex);
        }
    }

    private async Task<bool> ProcessFileAsync(
        string dolphinToolPath,
        string inputFile,
        string outputFolder,
        bool deleteOriginal,
        string compressionMethod,
        int compressionLevel,
        int blockSize,
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
                    compressionMethod,
                    compressionLevel,
                    blockSize,
                    cancellationToken);
            }
            else
            {
                return await ConvertSingleFileAsync(
                    dolphinToolPath,
                    inputFile,
                    outputFolder,
                    deleteOriginal,
                    compressionMethod,
                    compressionLevel,
                    blockSize,
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
            await _reportBugAsync($"Error processing file: {fileName}", ex);
            return false;
        }
    }

    private async Task<bool> ProcessArchiveFileAsync(
        string dolphinToolPath,
        string archivePath,
        string outputFolder,
        bool deleteOriginal,
        string compressionMethod,
        int compressionLevel,
        int blockSize,
        CancellationToken cancellationToken)
    {
        var archiveFileName = Path.GetFileName(archivePath);

        try
        {
            _logMessage($"Extracting archive: {archiveFileName}");

            var extractionResult = await ExtractArchiveAsync(archivePath, cancellationToken);
            if (!extractionResult.Success)
            {
                _logMessage($"Failed to extract {archiveFileName}: {extractionResult.ErrorMessage}");
                return false;
            }

            var extractedFilePath = extractionResult.FilePath;
            var tempDir = extractionResult.TempDir;
            var isRvzFile = extractionResult.IsRvzFile;

            try
            {
                bool success;

                // If the extracted file is already an RVZ, just copy it directly to the output folder
                if (isRvzFile)
                {
                    var fileName = Path.GetFileName(extractedFilePath);
                    var outputFile = Path.Combine(outputFolder, fileName);

                    _logMessage($"Found RVZ file inside archive, copying directly: {fileName}");

                    try
                    {
                        // Ensure the output directory exists
                        Directory.CreateDirectory(outputFolder);

                        // Copy the file
                        File.Copy(extractedFilePath, outputFile, true);
                        _logMessage($"Successfully copied RVZ file from archive: {fileName}");

                        success = true;
                    }
                    catch (Exception ex)
                    {
                        _logMessage($"Error copying RVZ file from archive {fileName}: {ex.Message}");
                        success = false;
                    }
                }
                else
                {
                    // Not an RVZ file, convert it as usual
                    success = await ConvertSingleFileAsync(
                        dolphinToolPath,
                        extractedFilePath,
                        outputFolder,
                        false, // Don't delete extracted file - we'll clean up temp dir
                        compressionMethod,
                        compressionLevel,
                        blockSize,
                        cancellationToken);
                }

                if (success && deleteOriginal)
                {
                    await TryDeleteFileAsync(archivePath, "original archive");
                }

                return success;
            }
            finally
            {
                await TryDeleteDirectoryAsync(tempDir, "temporary extraction directory");
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
            await _reportBugAsync($"Error processing archive: {archiveFileName}", ex);
            return false;
        }
    }

    private async Task<bool> ConvertSingleFileAsync(
        string dolphinToolPath,
        string inputFile,
        string outputFolder,
        bool deleteOriginal,
        string compressionMethod,
        int compressionLevel,
        int blockSize,
        CancellationToken cancellationToken)
    {
        var fileName = Path.GetFileName(inputFile);
        var outputFileName = _fileService.GetBaseFileNameWithoutGameExtension(fileName) + ".rvz";
        var outputFile = Path.Combine(outputFolder, outputFileName);

        try
        {
            // If the input file is already an RVZ file, just copy it to the output folder instead of converting
            if (_fileService.IsRvzFile(inputFile))
            {
                _logMessage($"File is already in RVZ format, copying: {fileName} -> {outputFileName}");

                try
                {
                    // Ensure the output directory exists
                    Directory.CreateDirectory(outputFolder);

                    // Copy the file
                    File.Copy(inputFile, outputFile, true);
                    _logMessage($"Successfully copied RVZ file: {fileName}");

                    // Delete original if requested
                    if (deleteOriginal)
                    {
                        await TryDeleteFileAsync(inputFile, "original file");
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    _logMessage($"Error copying RVZ file {fileName}: {ex.Message}");
                    return false;
                }
            }
            else
            {
                _logMessage($"Converting: {fileName} -> {outputFileName}");

                var success = await ConvertToRvzAsync(
                    dolphinToolPath,
                    inputFile,
                    outputFile,
                    compressionMethod,
                    compressionLevel,
                    blockSize,
                    cancellationToken);

                if (success && deleteOriginal)
                {
                    await TryDeleteFileAsync(inputFile, "original file");
                }

                return success;
            }
        }
        catch (OperationCanceledException)
        {
            _logMessage($"Conversion canceled: {fileName}");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Error converting {fileName}: {ex.Message}");
            await _reportBugAsync($"Error converting file: {fileName}", ex);
            return false;
        }
    }

    private async Task<bool> ConvertToRvzAsync(
        string dolphinToolPath,
        string inputFile,
        string outputFile,
        string compressionMethod,
        int compressionLevel,
        int blockSize,
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
            process.StartInfo.ArgumentList.Add("rvz");
            process.StartInfo.ArgumentList.Add("-c");
            process.StartInfo.ArgumentList.Add(compressionMethod);
            process.StartInfo.ArgumentList.Add("-l");
            process.StartInfo.ArgumentList.Add(compressionLevel.ToString(CultureInfo.InvariantCulture));
            process.StartInfo.ArgumentList.Add("-b");
            process.StartInfo.ArgumentList.Add(blockSize.ToString(CultureInfo.InvariantCulture));

            process.EnableRaisingEvents = true;

            // Bounded channel prevents unbounded memory growth if reader is slower than producer
            var outputChannel = Channel.CreateBounded<string>(new BoundedChannelOptions(1000)
            {
                FullMode = BoundedChannelFullMode.Wait
            });
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
                    outputChannel.Writer.TryWrite(args.Data);
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
                    outputChannel.Writer.TryWrite(args.Data);
                    _logMessage($"[DolphinTool ERROR] {args.Data}");
                }
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            await process.WaitForExitAsync(cancellationToken);
            await Task.WhenAll(outputCompleted.Task, errorCompleted.Task);

            outputChannel.Writer.Complete();

            var outputBuilder = new StringBuilder();
            await foreach (var line in outputChannel.Reader.ReadAllAsync(cancellationToken))
            {
                outputBuilder.AppendLine(line);
            }

            var output = outputBuilder.ToString();
            if (process.ExitCode == 0)
            {
                _logMessage($"Successfully converted to RVZ: {Path.GetFileName(inputFile)}");
                return true;
            }
            else
            {
                _logMessage($"Conversion failed for {Path.GetFileName(inputFile)}. Exit code: {process.ExitCode}");
                _logMessage($"Output: {output}");
                return false;
            }
        }
        catch (OperationCanceledException)
        {
            if (!process.HasExited) process.Kill(true);
            throw;
        }
    }

    private async Task<(bool Success, string FilePath, string TempDir, string ErrorMessage, bool IsRvzFile)> ExtractArchiveAsync(string archivePath, CancellationToken cancellationToken)
    {
        var tempDir = string.Empty;

        try
        {
            // Create temporary directory for extraction
            tempDir = Path.Combine(Path.GetTempPath(), "BatchConvertToRVZ_Extract_" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);

            _logMessage($"Extracting archive to temporary directory: {tempDir}");

            using var archive = ArchiveFactory.OpenArchive(archivePath);
            var supportedExtensions = _fileService.GetPrimaryTargetExtensionsInsideArchive();
            var rvzExtensions = _fileService.GetRvzExtensions();

            // First, check for RVZ files inside the archive
            var rvzEntry = archive.Entries.FirstOrDefault(e =>
                e is { IsDirectory: false, Key: not null } &&
                rvzExtensions.Any(ext => e.Key.EndsWith(ext, StringComparison.OrdinalIgnoreCase)));

            // If there's an RVZ file, prefer it over other formats
            var entry = rvzEntry ?? archive.Entries.FirstOrDefault(e =>
                e is { IsDirectory: false, Key: not null } &&
                supportedExtensions.Any(ext => e.Key.EndsWith(ext, StringComparison.OrdinalIgnoreCase)));

            if (entry == null)
            {
                var archiveName = Path.GetFileName(archivePath);
                return (false, string.Empty, string.Empty, $"No supported disc image found inside {archiveName}.", false);
            }

            // Determine if the extracted file is an RVZ
            var isRvzFile = rvzExtensions.Any(ext =>
                entry.Key?.EndsWith(ext, StringComparison.OrdinalIgnoreCase) == true);

            // Extract the file name from the entry key, handling potential directory separators
            var entryName = entry.Key?.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
            entryName = Path.GetFileName(entryName);

            // If entry name is still invalid, create one from archive name preserving the extension
            if (string.IsNullOrWhiteSpace(entryName))
            {
                var archiveName = Path.GetFileNameWithoutExtension(archivePath);
                var entryExtension = supportedExtensions.FirstOrDefault(ext =>
                    entry.Key?.EndsWith(ext, StringComparison.OrdinalIgnoreCase) == true) ?? ".iso";
                entryName = $"{archiveName}{entryExtension}";
            }

            var extractedFilePath = Path.Combine(tempDir, entryName);

            await using (var source = await entry.OpenEntryStreamAsync(cancellationToken).ConfigureAwait(false))
            await using (var destination = File.Create(extractedFilePath))
            {
                await source.CopyToAsync(destination, cancellationToken).ConfigureAwait(false);
            }

            if (isRvzFile)
            {
                _logMessage($"Extracted RVZ file {entryName} from archive.");
            }
            else
            {
                _logMessage($"Extracted {entryName} from archive.");
            }

            return (true, extractedFilePath, tempDir, string.Empty, isRvzFile);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            // Real user cancellation - propagate
            throw;
        }
        catch (Exception ex)
        {
            // SharpCompress internal failure (not user cancellation) - try 7za.exe fallback
            var archiveName = Path.GetFileName(archivePath);

            if (ex is OperationCanceledException)
            {
                _logMessage($"SharpCompress extraction failed for {archiveName} (internal cancellation), falling back to 7za.exe...");
            }
            else
            {
                _logMessage($"SharpCompress extraction failed for {archiveName}: {ex.Message}, falling back to 7za.exe...");
            }

            // Clean up SharpCompress temp dir
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

            // Try 7za.exe fallback
            var sevenZipResult = await ExtractWith7ZipAsync(archivePath, cancellationToken);
            if (sevenZipResult.Success)
            {
                return sevenZipResult;
            }

            // Both SharpCompress and 7za.exe failed - report as corrupt
            _logMessage($"Extraction failed with both SharpCompress and 7za.exe for {archiveName}. File may be corrupt.");
            await _reportBugAsync($"Error extracting archive: {archiveName}", ex);

            return (false, string.Empty, string.Empty, $"Failed to extract archive (file may be corrupt): {archiveName}", false);
        }
    }

    private async Task<(bool Success, string FilePath, string TempDir, string ErrorMessage, bool IsRvzFile)> ExtractWith7ZipAsync(string archivePath, CancellationToken cancellationToken)
    {
        var tempDir = string.Empty;

        try
        {
            tempDir = Path.Combine(Path.GetTempPath(), "BatchConvertToRVZ_7Zip_" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);

            var sevenZipPath = Get7ZipExecutablePath();
            if (!File.Exists(sevenZipPath))
            {
                _logMessage($"7za executable not found at: {sevenZipPath}");
                return (false, string.Empty, string.Empty, "7za executable not found.", false);
            }

            _logMessage($"Extracting with 7za.exe to: {tempDir}");

            using var process = new Process();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = sevenZipPath,
                Arguments = $"x -o\"{tempDir}\" -y \"{archivePath}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();

            process.OutputDataReceived += (_, args) =>
            {
                if (args.Data is not null) outputBuilder.AppendLine(args.Data);
            };
            process.ErrorDataReceived += (_, args) =>
            {
                if (args.Data is not null) errorBuilder.AppendLine(args.Data);
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            await process.WaitForExitAsync(cancellationToken);

            if (process.ExitCode != 0)
            {
                var errorOutput = errorBuilder.ToString();
                _logMessage($"7za.exe extraction failed with exit code {process.ExitCode}: {errorOutput}");
                return (false, string.Empty, string.Empty, $"7za.exe extraction failed: {errorOutput}", false);
            }

            // Find the extracted disc image
            var supportedExtensions = _fileService.GetPrimaryTargetExtensionsInsideArchive();
            var rvzExtensions = _fileService.GetRvzExtensions();

            var extractedFile = Directory.GetFiles(tempDir, "*", SearchOption.AllDirectories)
                .FirstOrDefault(f =>
                {
                    var ext = Path.GetExtension(f).ToLowerInvariant();
                    return supportedExtensions.Contains(ext) || rvzExtensions.Contains(ext);
                });

            if (extractedFile is null)
            {
                _logMessage("No supported disc image found in 7za.exe extraction output.");
                return (false, string.Empty, string.Empty, "No supported disc image found after 7za.exe extraction.", false);
            }

            var isRvzFile = rvzExtensions.Contains(Path.GetExtension(extractedFile).ToLowerInvariant());
            var entryName = Path.GetFileName(extractedFile);

            if (isRvzFile)
            {
                _logMessage($"Extracted RVZ file {entryName} from archive using 7za.exe.");
            }
            else
            {
                _logMessage($"Extracted {entryName} from archive using 7za.exe.");
            }

            return (true, extractedFile, tempDir, string.Empty, isRvzFile);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"7za.exe extraction error: {ex.Message}");

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

            return (false, string.Empty, string.Empty, $"7za.exe extraction error: {ex.Message}", false);
        }
    }

    private static string Get7ZipExecutablePath()
    {
        var architecture = RuntimeInformation.ProcessArchitecture;
        var exeName = architecture switch
        {
            Architecture.Arm64 => "7za_arm64.exe",
            _ => "7za.exe"
        };

        return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, exeName);
    }

    private Task<bool> TryDeleteFileAsync(string filePath, string description)
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

    private Task TryDeleteDirectoryAsync(string dirPath, string description)
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
