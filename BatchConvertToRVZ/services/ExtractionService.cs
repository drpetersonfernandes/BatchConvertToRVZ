using System.Diagnostics;
using System.IO;
using System.Text;

namespace BatchConvertToRVZ.services;

/// <summary>
/// Service responsible for extracting RVZ files to ISO format.
/// </summary>
public class ExtractionService
{
    private readonly Action<string> _logMessage;
    private readonly Func<string, Task> _reportBugAsync;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExtractionService"/> class.
    /// </summary>
    /// <param name="logMessage">Action to log messages.</param>
    /// <param name="reportBugAsync">Function to report bugs asynchronously.</param>
    public ExtractionService(
        Action<string> logMessage,
        Func<string, Task> reportBugAsync)
    {
        _logMessage = logMessage;
        _reportBugAsync = reportBugAsync;
    }

    /// <summary>
    /// Performs batch extraction of RVZ files to ISO format.
    /// </summary>
    /// <param name="dolphinToolPath">Path to the DolphinTool executable.</param>
    /// <param name="files">Array of file paths to extract.</param>
    /// <param name="outputFolder">Output folder for extracted files.</param>
    /// <param name="deleteFiles">Whether to delete original files after successful extraction.</param>
    /// <param name="updateProgress">Callback to update progress.</param>
    /// <param name="incrementSuccess">Callback to increment success count.</param>
    /// <param name="incrementFailure">Callback to increment failure count.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public async Task PerformBatchExtractionAsync(
        string dolphinToolPath,
        string[] files,
        string outputFolder,
        bool deleteFiles,
        Action<int, int, string> updateProgress,
        Action<int> incrementSuccess,
        Action<int> incrementFailure,
        CancellationToken cancellationToken)
    {
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

                var success = await ExtractSingleFileAsync(
                    dolphinToolPath,
                    inputFile,
                    outputFolder,
                    deleteFiles,
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
            await _reportBugAsync($"Error during batch extraction operation: {ex.Message}");
        }
    }

    private async Task<bool> ExtractSingleFileAsync(
        string dolphinToolPath,
        string inputFile,
        string outputFolder,
        bool deleteOriginal,
        CancellationToken cancellationToken)
    {
        var fileName = Path.GetFileName(inputFile);
        var fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
        var outputFileName = fileNameWithoutExt + ".iso";
        var outputFile = Path.Combine(outputFolder, outputFileName);

        try
        {
            _logMessage($"Converting to ISO: {fileName} -> {outputFileName}");

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

            var success = await ConvertToIsoAsync(
                dolphinToolPath,
                inputFile,
                outputFile,
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

    private async Task<bool> ConvertToIsoAsync(
        string dolphinToolPath,
        string inputFile,
        string outputFile,
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
            process.StartInfo.ArgumentList.Add("iso");

            process.EnableRaisingEvents = true;

            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();

            process.OutputDataReceived += (_, args) =>
            {
                if (args.Data != null)
                {
                    outputBuilder.AppendLine(args.Data);
                    _logMessage($"[DolphinTool] {args.Data}");
                }
            };

            process.ErrorDataReceived += (_, args) =>
            {
                if (args.Data != null)
                {
                    errorBuilder.AppendLine(args.Data);
                    _logMessage($"[DolphinTool ERROR] {args.Data}");
                }
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            await process.WaitForExitAsync(cancellationToken);

            var output = outputBuilder.ToString();
            var error = errorBuilder.ToString();

            // Standardize success check: exit code 0 AND success message in output AND file exists
            if (process.ExitCode == 0 && output.Contains("Successfully converted"))
            {
                // Wait for the file to be written (with timeout)
                var fileExists = await WaitForFileExistsAsync(outputFile, cancellationToken);

                if (fileExists)
                {
                    _logMessage($"Successfully converted to ISO: {Path.GetFileName(outputFile)}");
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
}
