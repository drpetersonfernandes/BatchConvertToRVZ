using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace BatchConvertToRVZ.services;

public class VerificationService
{
    private readonly Action<string> _logMessage;
    private readonly Func<string, Task> _reportBugAsync;

    public VerificationService(
        Action<string> logMessage,
        Func<string, Task> reportBugAsync)
    {
        _logMessage = logMessage;
        _reportBugAsync = reportBugAsync;
    }

    public async Task PerformBatchVerificationAsync(
        string dolphinToolPath,
        string[] files,
        bool moveFailed,
        bool moveSuccess,
        Action<int, int, string> updateProgress,
        Action<int> incrementSuccess,
        Action<int> incrementFailure,
        CancellationToken cancellationToken)
    {
        try
        {
            _logMessage("Preparing for batch verification...");

            var totalFilesToProcess = files.Length;
            _logMessage($"Verifying {totalFilesToProcess} selected RVZ files.");

            if (totalFilesToProcess == 0)
            {
                _logMessage("No files selected for verification.");
                return;
            }

            var filesProcessedCount = 0;

            _logMessage("Verifying files sequentially.");
            foreach (var inputFile in files)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var fileName = Path.GetFileName(inputFile);
                var baseFolder = Path.GetDirectoryName(inputFile) ?? string.Empty;
                var success = await VerifyRzvFileAsync(
                    dolphinToolPath,
                    inputFile,
                    baseFolder,
                    moveFailed,
                    moveSuccess,
                    cancellationToken);

                if (success)
                {
                    incrementSuccess(1);
                }
                else
                {
                    incrementFailure(1);
                }

                filesProcessedCount++;
                updateProgress(filesProcessedCount, totalFilesToProcess, fileName);
            }
        }
        catch (OperationCanceledException)
        {
            _logMessage("Batch verification operation was canceled.");
            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Error during batch verification: {ex.Message}");
            await _reportBugAsync($"Error during batch verification operation: {ex.Message}");
        }
    }

    private async Task<bool> VerifyRzvFileAsync(
        string dolphinToolPath,
        string inputFile,
        string baseFolder,
        bool moveFailed,
        bool moveSuccess,
        CancellationToken token)
    {
        var fileName = Path.GetFileName(inputFile);
        using var process = new Process();
        var verificationResult = false;
        string? tempWorkingDirectory = null;

        DataReceivedEventHandler? outputHandler = null;
        DataReceivedEventHandler? errorHandler = null;

        try
        {
            _logMessage($"Verifying: {fileName}...");

            tempWorkingDirectory = Path.Combine(Path.GetTempPath(), "BatchConvertToRVZ_DolphinTool_Temp_" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempWorkingDirectory);

            process.StartInfo = new ProcessStartInfo
            {
                FileName = dolphinToolPath,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = tempWorkingDirectory
            };
            process.StartInfo.ArgumentList.Add("verify");
            process.StartInfo.ArgumentList.Add("-i");
            process.StartInfo.ArgumentList.Add(inputFile);

            process.EnableRaisingEvents = true;

            var outputQueue = new System.Collections.Concurrent.ConcurrentQueue<string>();
            var outputCompleted = new TaskCompletionSource<bool>();
            var errorCompleted = new TaskCompletionSource<bool>();

            outputHandler = (_, args) =>
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
            errorHandler = (_, args) =>
            {
                if (args.Data is null)
                {
                    errorCompleted.TrySetResult(true);
                }
                else
                {
                    outputQueue.Enqueue(args.Data);
                    _logMessage($"[DolphinTool ERROR] {args.Data}");
                }
            };
            process.OutputDataReceived += outputHandler;
            process.ErrorDataReceived += errorHandler;

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            await process.WaitForExitAsync(token);
            await Task.WhenAll(outputCompleted.Task, errorCompleted.Task);

            var outputBuilder = new StringBuilder();
            while (outputQueue.TryDequeue(out var line)) outputBuilder.AppendLine(line);
            var output = outputBuilder.ToString();

            if (process.ExitCode == 0 && output.Contains("Problems Found: No"))
            {
                verificationResult = true;
                _logMessage($"Verification successful: {fileName}");

                if (moveSuccess)
                {
                    await MoveFileToSubfolder(inputFile, baseFolder, "_Success");
                }
            }
            else
            {
                _logMessage($"Verification failed: {fileName}");
                _logMessage($"Exit code: {process.ExitCode}");
                _logMessage($"Output: {output}");

                if (moveFailed)
                {
                    await MoveFileToSubfolder(inputFile, baseFolder, "_Failed");
                }
            }
        }
        catch (OperationCanceledException)
        {
            _logMessage($"Verification canceled: {fileName}");
            if (!process.HasExited)
            {
                process.Kill(true);
            }

            throw;
        }
        catch (Exception ex)
        {
            _logMessage($"Error verifying {fileName}: {ex.Message}");
            verificationResult = false;
        }
        finally
        {
            if (outputHandler != null)
            {
                process.OutputDataReceived -= outputHandler;
            }

            if (errorHandler != null)
            {
                process.ErrorDataReceived -= errorHandler;
            }

            try
            {
                if (tempWorkingDirectory != null && Directory.Exists(tempWorkingDirectory))
                {
                    Directory.Delete(tempWorkingDirectory, true);
                }
            }
            catch (Exception ex)
            {
                _logMessage($"Failed to clean up temp directory: {ex.Message}");
            }
        }

        return verificationResult;
    }

    private Task MoveFileToSubfolder(string sourceFilePath, string baseFolder, string subfolderName)
    {
        try
        {
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

            File.Move(sourceFilePath, destinationPath);
            _logMessage($"Moved {fileName} to {subfolderName} folder.");
        }
        catch (Exception ex)
        {
            _logMessage($"Failed to move file to {subfolderName} folder: {ex.Message}");
        }

        return Task.CompletedTask;
    }
}
