using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace BatchConvertToRVZ.services;

public class ConversionService
{
    private readonly Action<string> _logMessage;
    private readonly Func<string, Task> _reportBugAsync;

    // Supported input extensions
    private static readonly string[] ArchiveExtensions = [".zip", ".7z", ".rar"];

    public ConversionService(
        Action<string> logMessage,
        Func<string, Task> reportBugAsync)
    {
        _logMessage = logMessage;
        _reportBugAsync = reportBugAsync;
    }

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
                if (cancellationToken.IsCancellationRequested)
                {
                    _logMessage("Operation canceled by user.");
                    break;
                }

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
        }
        catch (Exception ex)
        {
            _logMessage($"Error during batch conversion: {ex.Message}");
            await _reportBugAsync($"Error during batch conversion operation: {ex.Message}");
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
            if (ArchiveExtensions.Contains(inputExtension))
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

            var extractionResult = await ExtractArchiveAsync();
            if (!extractionResult.Success)
            {
                _logMessage($"Failed to extract {archiveFileName}: {extractionResult.ErrorMessage}");
                return false;
            }

            var extractedFilePath = extractionResult.FilePath;
            var tempDir = extractionResult.TempDir;

            try
            {
                var success = await ConvertSingleFileAsync(
                    dolphinToolPath,
                    extractedFilePath,
                    outputFolder,
                    false, // Don't delete extracted file - we'll clean up temp dir
                    compressionMethod,
                    compressionLevel,
                    blockSize,
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
        var outputFileName = Path.ChangeExtension(fileName, ".rvz");
        var outputFile = Path.Combine(outputFolder, outputFileName);

        try
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

            var outputBuilder = new StringBuilder();
            process.OutputDataReceived += (_, args) =>
            {
                if (!string.IsNullOrEmpty(args.Data))
                {
                    outputBuilder.AppendLine(args.Data);
                    _logMessage($"[DolphinTool] {args.Data}");
                }
            };

            process.ErrorDataReceived += (_, args) =>
            {
                if (!string.IsNullOrEmpty(args.Data))
                {
                    outputBuilder.AppendLine(args.Data);
                    _logMessage($"[DolphinTool ERROR] {args.Data}");
                }
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            await process.WaitForExitAsync(cancellationToken);

            var output = outputBuilder.ToString();
            if (process.ExitCode == 0 && output.Contains("Successfully converted"))
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
        finally
        {
            if (!process.HasExited) process.Kill(true);
        }
    }

    private static Task<(bool Success, string FilePath, string TempDir, string ErrorMessage)> ExtractArchiveAsync()
    {
        try
        {
            // This is a simplified version - actual implementation would use SharpCompress
            // For now, return a stub implementation
            return Task.FromResult((false, string.Empty, string.Empty, "Extraction not implemented in ConversionService"));
        }
        catch (Exception exception)
        {
            return Task.FromException<(bool Success, string FilePath, string TempDir, string ErrorMessage)>(exception);
        }
    }

    private Task<bool> TryDeleteFile(string filePath, string description)
    {
        try
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
        catch (Exception exception)
        {
            return Task.FromException<bool>(exception);
        }
    }

    private Task TryDeleteDirectory(string dirPath, string description)
    {
        try
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
        catch (Exception exception)
        {
            return Task.FromException(exception);
        }
    }
}