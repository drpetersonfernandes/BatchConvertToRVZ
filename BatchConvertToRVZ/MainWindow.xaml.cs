﻿using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression; // Added for ZipFile
using System.Text;
using System.Windows;
using Microsoft.Win32;
using System.Text.RegularExpressions;

namespace BatchConvertToRVZ;

public partial class MainWindow : IDisposable
{
    private CancellationTokenSource _cts;
    private readonly BugReportService _bugReportService;

    // Bug Report API configuration
    private const string BugReportApiUrl = "https://www.purelogiccode.com/bugreport/api/send-bug-report";
    private const string BugReportApiKey = "hjh7yu6t56tyr540o9u8767676r5674534453235264c75b6t7ggghgg76trf564e";
    private const string ApplicationName = "BatchConvertToRVZ";

    private int _currentDegreeOfParallelismForFiles = 1;

    // DolphinTool specific constants
    private const string DolphinToolExeName = "DolphinTool.exe";
    private const string RvzCompressionMethod = "zstd"; // Default compression method
    private const int RvzCompressionLevel = 5; // Default compression level
    private const int RvzBlockSize = 131072; // Default block size

    // 7z specific constants and fields (Re-added)
    private const string SevenZipExeName = "7z.exe";
    private readonly string _sevenZipPath;
    private readonly bool _isSevenZipAvailable;

    // Supported input extensions (Updated to include archives)
    private static readonly string[] AllSupportedInputExtensions = { ".iso", ".zip", ".7z", ".rar" };

    private static readonly string[] ArchiveExtensions = { ".zip", ".7z", ".rar" };

    // Primary target extension inside archives for RVZ conversion
    private static readonly string[] PrimaryTargetExtensionsInsideArchive = { ".iso" };


    public MainWindow()
    {
        InitializeComponent();
        _cts = new CancellationTokenSource();

        // Initialize the bug report service
        _bugReportService = new BugReportService(BugReportApiUrl, BugReportApiKey, ApplicationName);

        LogMessage("Welcome to the Batch Convert to RVZ.");
        LogMessage("");
        LogMessage("This program will convert GameCube/Wii ISO files (.iso) to RVZ format.");
        LogMessage("It also supports extracting ISOs from ZIP, 7Z, and RAR archives."); // Updated welcome message
        LogMessage("");
        LogMessage("Please follow these steps:");
        LogMessage("1. Select the input folder containing ISO files or archives to convert"); // Updated step 1
        LogMessage("2. Select the output folder where RVZ files will be saved");
        LogMessage("3. Choose whether to delete original files after conversion");
        LogMessage("4. Optionally, enable parallel processing for faster conversion of multiple files");
        LogMessage("5. Click 'Start Conversion' to begin the process");
        LogMessage("");

        var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var dolphinToolPath = Path.Combine(appDirectory, DolphinToolExeName);

        if (File.Exists(dolphinToolPath))
        {
            LogMessage($"{DolphinToolExeName} found in the application directory.");
        }
        else
        {
            LogMessage($"WARNING: {DolphinToolExeName} not found in the application directory!");
            LogMessage($"Please ensure {DolphinToolExeName} is in the same folder as this application.");
            // FIX: Changed Task.Run lambda syntax
            Task.Run(() => ReportBugAsync($"{DolphinToolExeName} not found in the application directory. This will prevent the application from functioning correctly."));
        }

        // Re-add 7z check
        _sevenZipPath = Path.Combine(appDirectory, SevenZipExeName);
        if (File.Exists(_sevenZipPath))
        {
            _isSevenZipAvailable = true;
            LogMessage($"{SevenZipExeName} found. .7z and .rar extraction enabled.");
        }
        else
        {
            _isSevenZipAvailable = false;
            LogMessage($"WARNING: {SevenZipExeName} not found. .7z and .rar extraction will be disabled.");
        }
    }

    private void Window_Closing(object sender, CancelEventArgs e)
    {
        // Signal cancellation to any ongoing background tasks
        _cts.Cancel();

        // Allow the window to close normally.
        // WPF will handle the rest of the shutdown process,
        // including disposing of window resources.
        // The Dispose method on MainWindow will be called by the framework
        // when the window's HwndSource is disposed during shutdown.
    }

    private void LogMessage(string message)
    {
        var timestampedMessage = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";

        // Use Dispatcher.BeginInvoke for potentially faster non-blocking updates
        Application.Current.Dispatcher.BeginInvoke(() =>
        {
            LogViewer.AppendText($"{timestampedMessage}{Environment.NewLine}");
            LogViewer.ScrollToEnd();
        });
    }

    private void BrowseInputButton_Click(object sender, RoutedEventArgs e)
    {
        // Updated description
        var inputFolder = SelectFolder("Select the folder containing ISO files or archives to convert");
        if (string.IsNullOrEmpty(inputFolder)) return;

        InputFolderTextBox.Text = inputFolder;
        LogMessage($"Input folder selected: {inputFolder}");
    }

    private void BrowseOutputButton_Click(object sender, RoutedEventArgs e)
    {
        var outputFolder = SelectFolder("Select the output folder where RVZ files will be saved");
        if (string.IsNullOrEmpty(outputFolder)) return;

        OutputFolderTextBox.Text = outputFolder;
        LogMessage($"Output folder selected: {outputFolder}");
    }

    private async void StartButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var dolphinToolPath = Path.Combine(appDirectory, DolphinToolExeName);

            if (!File.Exists(dolphinToolPath))
            {
                LogMessage($"Error: {DolphinToolExeName} not found in the application folder.");
                ShowError($"{DolphinToolExeName} is missing from the application folder. Please ensure it's in the same directory as this application.");
                await ReportBugAsync($"{DolphinToolExeName} not found when trying to start conversion",
                    new FileNotFoundException($"The required {DolphinToolExeName} file was not found.", dolphinToolPath));
                return;
            }

            var inputFolder = InputFolderTextBox.Text;
            var outputFolder = OutputFolderTextBox.Text;
            var deleteFiles = DeleteFilesCheckBox.IsChecked ?? false;
            var useParallelFileProcessing = ParallelProcessingCheckBox.IsChecked ?? false;

            if (useParallelFileProcessing)
            {
                _currentDegreeOfParallelismForFiles = 3; // Fixed maximum concurrency = 3
            }
            else
            {
                _currentDegreeOfParallelismForFiles = 1;
            }

            if (string.IsNullOrEmpty(inputFolder))
            {
                LogMessage("Error: No input folder selected.");
                ShowError("Please select the input folder containing files to convert.");
                return;
            }

            if (string.IsNullOrEmpty(outputFolder))
            {
                LogMessage("Error: No output folder selected.");
                ShowError("Please select the output folder where RVZ files will be saved.");
                return;
            }

            // Ensure output folder exists
            try
            {
                Directory.CreateDirectory(outputFolder);
            }
            catch (Exception ex)
            {
                LogMessage($"Error creating output directory {outputFolder}: {ex.Message}");
                ShowError($"Error creating output directory: {ex.Message}");
                await ReportBugAsync($"Error creating output directory: {outputFolder}", ex);
                return;
            }

            if (_cts.IsCancellationRequested)
            {
                _cts.Dispose();
                _cts = new CancellationTokenSource();
            }

            ClearProgressDisplay();
            SetControlsState(false);

            LogMessage("Starting batch conversion process...");
            LogMessage($"Using {DolphinToolExeName}: {dolphinToolPath}");
            LogMessage($"Input folder: {inputFolder}");
            LogMessage($"Output folder: {outputFolder}");
            LogMessage($"Delete original files: {deleteFiles}");
            LogMessage($"Parallel file processing: {useParallelFileProcessing} (Max concurrency: {_currentDegreeOfParallelismForFiles})");
            LogMessage($"RVZ Compression: Method={RvzCompressionMethod}, Level={RvzCompressionLevel}, Block Size={RvzBlockSize}");

            try
            {
                await PerformBatchConversionAsync(dolphinToolPath, inputFolder, outputFolder, deleteFiles, useParallelFileProcessing, _currentDegreeOfParallelismForFiles);
            }
            catch (OperationCanceledException)
            {
                LogMessage("Operation was canceled by user.");
            }
            catch (Exception ex)
            {
                LogMessage($"Error: {ex.Message}");
                await ReportBugAsync("Error during batch conversion process", ex);
            }
            finally
            {
                SetControlsState(true);
            }
        }
        catch (Exception ex)
        {
            await ReportBugAsync("Error during StartButton_Click", ex);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        _cts.Cancel();
        LogMessage("Cancellation requested. Waiting for current operation(s) to complete...");
    }

    private void SetControlsState(bool enabled)
    {
        InputFolderTextBox.IsEnabled = enabled;
        OutputFolderTextBox.IsEnabled = enabled;
        BrowseInputButton.IsEnabled = enabled;
        BrowseOutputButton.IsEnabled = enabled;
        DeleteFilesCheckBox.IsEnabled = enabled;
        ParallelProcessingCheckBox.IsEnabled = enabled;
        StartButton.IsEnabled = enabled;

        // Progress controls visibility is the inverse of enabled state
        ProgressBar.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        CancelButton.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        ProgressText.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;

        if (enabled)
        {
            ClearProgressDisplay();
        }
    }

    private static string? SelectFolder(string description)
    {
        var dialog = new OpenFolderDialog
        {
            Title = description
        };
        return dialog.ShowDialog() == true ? dialog.FolderName : null;
    }

    private async Task PerformBatchConversionAsync(string dolphinToolPath, string inputFolder, string outputFolder, bool deleteFiles, bool useParallelFileProcessing, int maxConcurrency)
    {
        try
        {
            LogMessage("Preparing for batch conversion...");

            // Find supported files (.iso, .zip, .7z, .rar) in the input folder
            var files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(file => AllSupportedInputExtensions.Contains(Path.GetExtension(file).ToLowerInvariant()))
                .ToArray();

            LogMessage($"Found {files.Length} files to process.");
            if (files.Length == 0)
            {
                LogMessage("No supported files (.iso, .zip, .7z, .rar) found in the input folder."); // Updated message
                return;
            }

            ProgressBar.Maximum = files.Length;
            ProgressBar.Value = 0;

            var successCount = 0;
            var failureCount = 0;
            var filesProcessedCount = 0;

            if (useParallelFileProcessing && files.Length > 1)
            {
                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = maxConcurrency,
                    CancellationToken = _cts.Token
                };

                await Parallel.ForEachAsync(files, parallelOptions, async (inputFile, token) =>
                {
                    var fileName = Path.GetFileName(inputFile);
                    LogMessage($"[Parallel] Starting: {fileName}");

                    var success = await ProcessFileAsync(dolphinToolPath, inputFile, outputFolder, deleteFiles);

                    if (success)
                    {
                        Interlocked.Increment(ref successCount);
                        LogMessage($"[Parallel] Successful: {fileName}");
                    }
                    else
                    {
                        Interlocked.Increment(ref failureCount);
                        LogMessage($"[Parallel] Failed: {fileName}");
                    }

                    var processed = Interlocked.Increment(ref filesProcessedCount);
                    var percentage = (double)processed / files.Length * 100;

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        ProgressBar.Value = processed;
                        ProgressText.Text = $"Processing file {processed} of {files.Length}: {fileName} ({percentage:F1}%)";
                        ProgressBar.Visibility = Visibility.Visible;
                        ProgressText.Visibility = Visibility.Visible;
                    });
                });
            }
            else
            {
                LogMessage("Using sequential processing.");
                foreach (var t in files)
                {
                    if (_cts.Token.IsCancellationRequested)
                    {
                        LogMessage("Operation canceled by user.");
                        break;
                    }

                    var fileName = Path.GetFileName(t);

                    LogMessage($"[Sequential] Processing: {fileName}");

                    var success = await ProcessFileAsync(dolphinToolPath, t, outputFolder, deleteFiles);

                    if (success)
                    {
                        successCount++;
                        LogMessage($"Conversion successful: {fileName}");
                    }
                    else
                    {
                        failureCount++;
                        LogMessage($"Conversion failed: {fileName}");
                    }

                    var processed = ++filesProcessedCount;
                    var percentage = (double)processed / files.Length * 100;

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        ProgressBar.Value = processed;
                        ProgressText.Text = $"Processing file {processed} of {files.Length}: {fileName} ({percentage:F1}%)";
                        ProgressBar.Visibility = Visibility.Visible;
                        ProgressText.Visibility = Visibility.Visible;
                    });
                }
            }

            LogMessage("");
            LogMessage("Batch conversion completed.");
            LogMessage($"Successfully converted: {successCount} files");
            if (failureCount > 0) LogMessage($"Failed to convert: {failureCount} files");

            ShowMessageBox($"Batch conversion completed.\n\nSuccessfully converted: {successCount} files\nFailed to convert: {failureCount} files",
                "Conversion Complete", MessageBoxButton.OK, failureCount > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
        }
        catch (OperationCanceledException)
        {
            LogMessage("Batch conversion operation was canceled.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error during batch conversion: {ex.Message}");
            ShowError($"Error during batch conversion: {ex.Message}");
            await ReportBugAsync("Error during batch conversion operation", ex);
        }
    }

    // Modified ProcessFileAsync to handle archives
    private async Task<bool> ProcessFileAsync(string dolphinToolPath, string inputFile, string outputFolder, bool deleteOriginal)
    {
        var fileToProcess = inputFile;
        var isArchiveFile = false;
        var tempDir = string.Empty;
        var fileExtension = Path.GetExtension(inputFile).ToLowerInvariant();

        try
        {
            if (ArchiveExtensions.Contains(fileExtension))
            {
                LogMessage($"Processing archive: {Path.GetFileName(inputFile)}");
                var extractResult = await ExtractArchiveAsync(inputFile);
                if (extractResult.Success)
                {
                    fileToProcess = extractResult.FilePath;
                    tempDir = extractResult.TempDir;
                    isArchiveFile = true;
                    LogMessage($"Using extracted file: {Path.GetFileName(fileToProcess)} from archive {Path.GetFileName(inputFile)}");
                }
                else
                {
                    LogMessage($"Error extracting archive {Path.GetFileName(inputFile)}: {extractResult.ErrorMessage}");
                    // Report extraction failure as a bug
                    await ReportBugAsync($"Error extracting archive: {Path.GetFileName(inputFile)}", new Exception(extractResult.ErrorMessage));
                    return false;
                }
            }

            // Now process the fileToProcess (either original ISO or extracted ISO)
            var outputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileToProcess) + ".rvz");

            // Check if output file already exists
            if (File.Exists(outputFile))
            {
                LogMessage($"Output file already exists, skipping: {Path.GetFileName(outputFile)}");
                return true; // Consider existing file as successful conversion for this run
            }

            var success = await ConvertToRvzAsync(dolphinToolPath, fileToProcess, outputFile);

            if (!success || !deleteOriginal) return success;

            if (isArchiveFile)
            {
                // Delete the original archive file
                TryDeleteFile(inputFile, $"original archive file: {Path.GetFileName(inputFile)}");
            }
            else
            {
                // Delete the original standalone ISO file
                TryDeleteFile(inputFile, $"original ISO file: {Path.GetFileName(inputFile)}");
            }

            return success;
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Processing cancelled for {Path.GetFileName(inputFile)}.");
            // Attempt to clean up partially created output file on cancellation
            var potentialOutputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileToProcess) + ".rvz");
            if (File.Exists(potentialOutputFile))
            {
                TryDeleteFile(potentialOutputFile, "partially created RVZ file after cancellation");
            }

            throw; // Re-throw cancellation exception
        }
        catch (Exception ex)
        {
            LogMessage($"Error processing file {Path.GetFileName(inputFile)}: {ex.Message}");
            await ReportBugAsync($"Error processing file: {Path.GetFileName(inputFile)}", ex);
            // Attempt to clean up partially created output file on error
            var potentialOutputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileToProcess) + ".rvz");
            if (File.Exists(potentialOutputFile))
            {
                TryDeleteFile(potentialOutputFile, "partially created RVZ file after error");
            }

            return false;
        }
        finally
        {
            // Clean up the temporary directory if one was created
            if (isArchiveFile && !string.IsNullOrEmpty(tempDir) && Directory.Exists(tempDir))
            {
                TryDeleteDirectory(tempDir, "temporary extraction directory");
            }
        }
    }

    private async Task<bool> ConvertToRvzAsync(string dolphinToolPath, string inputFile, string outputFile)
    {
        try
        {
            LogMessage($"Converting '{Path.GetFileName(inputFile)}' to '{Path.GetFileName(outputFile)}'...");

            // DolphinTool.exe convert -i input.iso -o output.rvz -f rvz -c compression -l compression_level -b block_size
            var arguments = $"convert -i \"{inputFile}\" -o \"{outputFile}\" -f rvz -c {RvzCompressionMethod} -l {RvzCompressionLevel} -b {RvzBlockSize}";

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = dolphinToolPath,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                },
                EnableRaisingEvents = true
            };

            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();

            // Capture output and error streams
            process.OutputDataReceived += (_, args) =>
            {
                if (string.IsNullOrEmpty(args.Data)) return;

                outputBuilder.AppendLine(args.Data);
                // Attempt to parse progress. If not handled, log the raw line.
                if (!UpdateConversionProgress(args.Data))
                {
                    LogMessage($"[DolphinTool] {args.Data}");
                }
            };
            process.ErrorDataReceived += (_, args) =>
            {
                if (string.IsNullOrEmpty(args.Data)) return;

                errorBuilder.AppendLine(args.Data);
                // Attempt to parse progress. If not handled, log the raw line.
                if (!UpdateConversionProgress(args.Data))
                {
                    // Log errors differently
                    LogMessage($"[DolphinTool ERROR] {args.Data}");
                }
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            await process.WaitForExitAsync(_cts.Token);

            LogMessage($"DolphinTool raw output for {Path.GetFileName(inputFile)}: {outputBuilder}");
            if (errorBuilder.Length > 0 && process.ExitCode != 0) LogMessage($"DolphinTool raw error for {Path.GetFileName(inputFile)}: {errorBuilder}");

            return process.ExitCode == 0;
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Conversion cancelled for {Path.GetFileName(inputFile)}.");
            // Attempt to clean up partially created output file on cancellation
            if (File.Exists(outputFile))
            {
                TryDeleteFile(outputFile, "partially created RVZ file after cancellation");
            }

            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error converting file {Path.GetFileName(inputFile)}: {ex.Message}");
            await ReportBugAsync($"Error converting file: {Path.GetFileName(inputFile)}", ex);
            // Attempt to clean up partially created output file on error
            if (File.Exists(outputFile))
            {
                TryDeleteFile(outputFile, "partially created RVZ file after error");
            }

            return false;
        }
    }

    private void TryDeleteFile(string filePath, string description)
    {
        try
        {
            if (!File.Exists(filePath)) return;

            // Ensure the file is not read-only before attempting deletion
            var attributes = File.GetAttributes(filePath);
            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                File.SetAttributes(filePath, attributes & ~FileAttributes.ReadOnly);
            }

            File.Delete(filePath);
            LogMessage($"Deleted {description}: {Path.GetFileName(filePath)}");
        }
        catch (Exception ex)
        {
            LogMessage($"Failed to delete {description} {Path.GetFileName(filePath)}: {ex.Message}");
            // FIX: Changed Task.Run lambda syntax
            Task.Run(() => ReportBugAsync($"Failed to delete {description}: {Path.GetFileName(filePath)}", ex));
        }
    }

    // Re-add TryDeleteDirectory
    private void TryDeleteDirectory(string dirPath, string description)
    {
        try
        {
            if (!Directory.Exists(dirPath)) return;

            // Use a loop with delay for robustness in case of file locks
            for (var i = 0; i < 5; i++)
            {
                try
                {
                    Directory.Delete(dirPath, true);
                    LogMessage($"Cleaned up {description}: {dirPath}");
                    return; // Success
                }
                catch (IOException)
                {
                    // Directory might still be in use, wait and retry
                    Thread.Sleep(50);
                }
            }

            // If loop finishes without success
            LogMessage($"Failed to clean up {description} {dirPath} after multiple retries.");
        }
        catch (Exception ex)
        {
            LogMessage($"Failed to clean up {description} {dirPath}: {ex.Message}");
        }
    }

    // Re-add ExtractArchiveAsync
    private async Task<(bool Success, string FilePath, string TempDir, string ErrorMessage)> ExtractArchiveAsync(string archivePath)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        var extension = Path.GetExtension(archivePath).ToLowerInvariant();
        var archiveFileName = Path.GetFileName(archivePath);

        try
        {
            Directory.CreateDirectory(tempDir);
            LogMessage($"Extracting {archiveFileName} to temporary directory: {tempDir}");

            switch (extension)
            {
                case ".zip":
                    await Task.Run(() => ZipFile.ExtractToDirectory(archivePath, tempDir, true), _cts.Token);
                    break;
                case ".7z" or ".rar" when !_isSevenZipAvailable:
                    return (false, string.Empty, tempDir, $"{SevenZipExeName} not found. Cannot extract {extension} files.");
                case ".7z" or ".rar":
                {
                    var process = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = _sevenZipPath,
                            Arguments = $"x \"{archivePath}\" -o\"{tempDir}\" -y",
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };
                    var outputBuilder = new StringBuilder();
                    var errorBuilder = new StringBuilder();
                    process.OutputDataReceived += (_, args) =>
                    {
                        if (args.Data != null) outputBuilder.AppendLine(args.Data);
                    };
                    process.ErrorDataReceived += (_, args) =>
                    {
                        if (args.Data != null) errorBuilder.AppendLine(args.Data);
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    try
                    {
                        await process.WaitForExitAsync(_cts.Token);
                    }
                    catch (OperationCanceledException)
                    {
                        if (!process.HasExited)
                        {
                            try
                            {
                                process.Kill(true);
                            }
                            catch
                            {
                                /* Ignore */
                            }
                        }

                        LogMessage($"Extraction of {archiveFileName} cancelled.");
                        throw; // Re-throw cancellation
                    }

                    if (process.ExitCode != 0)
                    {
                        LogMessage($"Error extracting {archiveFileName} with {SevenZipExeName}. Exit code: {process.ExitCode}. Output: {outputBuilder}. Error: {errorBuilder}");
                        return (false, string.Empty, tempDir, $"7z.exe failed. Error: {errorBuilder.ToString().Trim()}");
                    }

                    LogMessage($"{SevenZipExeName} output for {archiveFileName}: {outputBuilder}");
                    break;
                }
                default:
                    // This case should ideally not be hit due to the initial file filter,
                    // but included for completeness.
                    return (false, string.Empty, tempDir, $"Unsupported archive type: {extension}");
            }

            // Find the first supported primary file (.iso) in the extracted directory
            var supportedFile = Directory.GetFiles(tempDir, "*.*", SearchOption.AllDirectories)
                .FirstOrDefault(f => PrimaryTargetExtensionsInsideArchive.Contains(Path.GetExtension(f).ToLowerInvariant()));

            if (supportedFile != null)
            {
                return (true, supportedFile, tempDir, string.Empty);
            }

            return (false, string.Empty, tempDir, "No supported primary files (.iso) found in archive."); // Updated message
        }
        catch (OperationCanceledException)
        {
            // Clean up temp dir on cancellation
            TryDeleteDirectory(tempDir, $"cancelled extraction directory for {archiveFileName}");
            throw; // Re-throw cancellation
        }
        catch (Exception ex)
        {
            LogMessage($"Error extracting archive {archiveFileName}: {ex.Message}");
            await ReportBugAsync($"Error extracting archive: {archiveFileName}", ex);
            // Clean up temp dir on error
            TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");
            return (false, string.Empty, tempDir, $"Exception during extraction: {ex.Message}");
        }
    }


    // DeleteOriginalFilesAsync is no longer needed as deletion logic is in ProcessFileAsync


    /// <summary>
    /// Attempts to parse and log progress from a DolphinTool output line.
    /// </summary>
    /// <param name="progressLine">The line from DolphinTool's output.</param>
    /// <returns>True if the line was recognized and handled as a progress line, false otherwise.</returns>
    private bool UpdateConversionProgress(string progressLine)
    {
        try
        {
            // Attempt to parse percentage from the line
            // Regex looks for digits, optional comma/period, optional digits, followed by %
            var match = Regex.Match(progressLine, @"(\d+[\.,]?\d*)%");
            if (!match.Success) return false;

            var percentageStr = match.Groups[1].Value.Replace(',', '.');
            if (!double.TryParse(percentageStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var percentage))
                return false;
            // Use the percentage to log a specific progress message
            // This uses the 'percentage' variable, resolving the warning.
            // Note: DolphinTool output format might vary, this is based on a common pattern.
            // We are not using this to update the main progress bar, which tracks files.
            LogMessage($"DolphinTool Converting: {percentage:F1}%");
            return true; // Handled as a progress line

            // If we reach here, it wasn't a recognized progress line
        }
        catch (Exception ex)
        {
            LogMessage($"Error parsing DolphinTool progress line '{progressLine}': {ex.Message}");
            return false; // Failed to parse/handle
        }
    }

    // Helper methods for UI interaction (Re-added/Ensured Private)
    private void ShowMessageBox(string message, string title, MessageBoxButton buttons, MessageBoxImage icon)
    {
        // Ensure this runs on the UI thread
        Application.Current.Dispatcher.Invoke(() => MessageBox.Show(this, message, title, buttons, icon));
    }

    private void ShowError(string message)
    {
        ShowMessageBox(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private async Task ReportBugAsync(string message, Exception? exception = null)
    {
        try
        {
            var fullReport = new StringBuilder();
            fullReport.AppendLine("=== Bug Report ===");
            fullReport.AppendLine($"Application: {ApplicationName}"); // Use the constant
            fullReport.AppendLine(CultureInfo.InvariantCulture, $"Version: {GetType().Assembly.GetName().Version}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $"OS: {Environment.OSVersion}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $".NET Version: {Environment.Version}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $"Date/Time: {DateTime.Now}");
            fullReport.AppendLine();
            fullReport.AppendLine("=== Error Message ===");
            fullReport.AppendLine(message);
            fullReport.AppendLine();

            if (exception != null)
            {
                fullReport.AppendLine("=== Exception Details ===");
                AppendExceptionDetailsToReport(fullReport, exception);
            }

            // Safely get log content from UI thread
            if (LogViewer != null)
            {
                var logContent = string.Empty;
                await Application.Current.Dispatcher.InvokeAsync(() => logContent = LogViewer.Text);
                if (!string.IsNullOrEmpty(logContent))
                {
                    fullReport.AppendLine().AppendLine("=== Application Log ===").Append(logContent);
                }
            }

            await _bugReportService.SendBugReportAsync(fullReport.ToString());
        }
        catch
        {
            /* Silently fail reporting */
        }
    }

    private static void AppendExceptionDetailsToReport(StringBuilder sb, Exception? ex, int level = 0)
    {
        while (ex != null)
        {
            var indent = new string(' ', level * 2);
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Type: {ex.GetType().FullName}");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Message: {ex.Message}");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Source: {ex.Source}");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}StackTrace:");
            sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}{ex.StackTrace}");
            if (ex.InnerException != null)
            {
                sb.AppendLine(CultureInfo.InvariantCulture, $"{indent}Inner Exception:");
                ex = ex.InnerException;
                level++;
            }
            else
            {
                break;
            }
        }
    }


    private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            new AboutWindow { Owner = this }.ShowDialog();
        }
        catch (Exception ex)
        {
            LogMessage($"Error opening About window: {ex.Message}");
            // FIX: Changed Task.Run lambda syntax
            Task.Run(() => ReportBugAsync("Error opening About window", ex));
        }
    }

    private void ClearProgressDisplay()
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            ProgressBar.Visibility = Visibility.Collapsed;
            ProgressText.Text = string.Empty;
            ProgressText.Visibility = Visibility.Collapsed;
            ProgressBar.Value = 0;
        });
    }

    public void Dispose()
    {
        _cts?.Cancel();
        _cts?.Dispose();
        _bugReportService?.Dispose();
        GC.SuppressFinalize(this);
    }
}