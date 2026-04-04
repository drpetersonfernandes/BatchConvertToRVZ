using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using SharpCompress.Archives;
using SharpCompress.Readers;
using SharpCompress.Common;

namespace BatchConvertToRVZ;

public partial class MainWindow : IDisposable
{
    private bool _disposed;
    private bool _isClosing;
    private Task? _runningTask;
    private bool _dependenciesOk;
    private string? _dolphinToolPath;
    private CancellationTokenSource _cts;
    private readonly services.UpdateService _updateService;

    private const string GitHubApiUrl = "https://api.github.com/repos/drpetersonfernandes/BatchConvertToRVZ/releases/latest";

    private int _currentDegreeOfParallelismForFiles = 1;

    // Compression settings (now instance variables to allow user configuration)
    private string _rvzCompressionMethod = "zstd"; // Default compression method
    private int _rvzCompressionLevel = 5; // Default compression level
    private int _rvzBlockSize = 131072; // Default block size (128KB)

    // Compression level ranges for different methods
    private static readonly Dictionary<string, (int Min, int Max)> CompressionLevelRanges = new()
    {
        { "zstd", (1, 22) },
        { "zlib", (1, 9) },
        { "lzma", (1, 9) },
        { "lzma2", (1, 9) },
        { "bzip2", (1, 9) },
        { "lz4", (1, 12) }
    };

    // Supported input extensions (Updated to include archives)
    private static readonly string[] AllSupportedInputExtensions = [".iso", ".zip", ".7z", ".rar"];

    private static readonly string[] ArchiveExtensions = [".zip", ".7z", ".rar"];

    // Primary target extensions inside archives for RVZ conversion
    private static readonly string[] PrimaryTargetExtensionsInsideArchive = [".iso", ".gcm", ".wbfs", ".nkit.iso"];

    // Supported extension for verification
    private static readonly string[] RvzExtension = [".rvz"];

    // Pre-formatted extension lists for user-facing messages
    private static readonly string SupportedInputExtensionsDisplay = string.Join(", ", AllSupportedInputExtensions);
    private static readonly string PrimaryTargetExtensionsDisplay = string.Join(", ", PrimaryTargetExtensionsInsideArchive);
    private static readonly string RvzExtensionDisplay = string.Join(", ", RvzExtension);

    // Statistics
    private int _totalFilesToProcess;
    private int _successCount;
    private int _failureCount;
    private readonly Stopwatch _operationTimer = new();

    // For Write Speed Calculation
    private const int WriteSpeedUpdateIntervalMs = 1000;

    // Thread-safe fields for aggregate write speed tracking during parallel processing
    private long _aggregateTotalBytesWritten;
    private DateTime _aggregateLastCheckTime = DateTime.UtcNow;
    private readonly object _aggregateSpeedLock = new();
    private int _activeConversionCount;

    // Fields for verification move options
    private bool _moveFailedFiles;
    private bool _moveSuccessFiles;

    // Counter to track active extractions (supports concurrent extractions in parallel mode)
    private int _activeExtractionCount;

    // Log batching
    private readonly System.Threading.Channels.Channel<string> _logChannel = System.Threading.Channels.Channel.CreateUnbounded<string>();
    private readonly Task? _logProcessorTask;

    public MainWindow()
    {
        InitializeComponent();
        _cts = new CancellationTokenSource();

        // Start log processor
        _logProcessorTask = Task.Run(ProcessLogsAsync);

        // The BugReportService is now initialized and managed by the App class.
        _updateService = new services.UpdateService(GitHubApiUrl);

        LogMessage("Welcome to the Batch Convert to RVZ.");
        LogMessage("");
        LogMessage("This program will convert GameCube/Wii ISO files (.iso) to RVZ format.");
        LogMessage("It also supports extracting ISOs from ZIP, 7Z, and RAR archives.");
        LogMessage("");
        LogMessage("Use the 'Convert' tab to convert ISO files or archives to RVZ.");
        LogMessage("Use the 'Verify Integrity' tab to check your existing RVZ files.");
        LogMessage("");

        CheckDependencies();

        ResetOperationStats();
        Loaded += MainWindow_Loaded;
    }

    private void CheckDependencies()
    {
        var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var missingFiles = new List<string>();
        string? dolphinToolExeName;

        try
        {
            dolphinToolExeName = GetDolphinToolExecutableName();
            _dolphinToolPath = Path.Combine(appDirectory, dolphinToolExeName);
            if (!File.Exists(_dolphinToolPath)) missingFiles.Add(dolphinToolExeName);
        }
        catch (PlatformNotSupportedException ex)
        {
            var errorMessage = $"Unsupported platform architecture. {ex.Message}";
            LogMessage($"ERROR: {errorMessage}");
            ShowError(errorMessage);
            Task.Run(() => ReportBugAsync($"Unsupported platform: {ex.Message}", ex));
            _dependenciesOk = false;
            StartConversionButton.IsEnabled = false;
            StartVerifyButton.IsEnabled = false;
            return; // Stop further checks
        }

        if (missingFiles.Count != 0)
        {
            _dependenciesOk = false;
            var missingFilesString = string.Join(", ", missingFiles);
            var errorMessage = $"The following critical file(s) are missing: {missingFilesString}.\n\nThe application cannot function without them. Please ensure all files from the release archive are in the same folder as this application.";
            LogMessage($"WARNING: {errorMessage.ReplaceLineEndings(" ")}");
            ShowError(errorMessage);
            Task.Run(() => ReportBugAsync($"Missing critical files: {missingFilesString}"));
        }
        else
        {
            _dependenciesOk = true;
            StartConversionButton.IsEnabled = true;
            StartVerifyButton.IsEnabled = true;
            LogMessage($"{dolphinToolExeName} found in the application directory.");
            LogMessage("SharpCompress library loaded for archive extraction.");
        }

        LogMessage("");
    }

    private static string GetDolphinToolExecutableName()
    {
        var architecture = RuntimeInformation.ProcessArchitecture;
        // ReSharper disable once SwitchExpressionHandlesSomeKnownEnumValuesWithExceptionInDefault
        return architecture switch
        {
            Architecture.X64 => "DolphinTool.exe",
            Architecture.Arm64 => "DolphinTool_arm64.exe",
            _ => throw new PlatformNotSupportedException($"Unsupported architecture: {architecture}")
        };
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        try
        {
            // Check for updates on startup in the background
            await CheckForUpdatesAsync(false);
        }
        catch (Exception ex)
        {
            _ = ReportBugAsync("Error during startup", ex);
        }
    }

    private async void Window_Closing(object? sender, CancelEventArgs e)
    {
        try
        {
            if (_isClosing)
                return;

            if (_runningTask is not { IsCompleted: false })
            {
                // Ensure log processor is shut down
                _logChannel.Writer.Complete();
                if (_logProcessorTask != null) await _logProcessorTask;
                return;
            }

            e.Cancel = true;
            _isClosing = true;

            try
            {
                await _cts.CancelAsync();

                // Wait for the task to complete with a timeout to prevent hanging the app on exit
                var timeoutTask = Task.Delay(5000); // 5 seconds timeout
                var completedTask = await Task.WhenAny(_runningTask, timeoutTask);

                if (completedTask == timeoutTask)
                {
                    LogMessage("Warning: Tasks did not cancel within the timeout period. Forcing exit.");
                }
            }
            catch (Exception ex)
            {
                _isClosing = false;
                _ = ReportBugAsync("Error during task cancellation while closing", ex);
                // Even on error, we should try to close if we can't cancel
            }

            // Complete the log channel and wait for processor to finish
            _logChannel.Writer.Complete();
            if (_logProcessorTask != null) await Task.WhenAny(_logProcessorTask, Task.Delay(1000));

            try
            {
                if (Application.Current?.Dispatcher is { HasShutdownStarted: false, HasShutdownFinished: false } dispatcher)
                {
                    _ = dispatcher.BeginInvoke(() =>
                    {
                        if (!_disposed)
                            Close();
                    }, System.Windows.Threading.DispatcherPriority.Normal);
                }
            }
            catch (Exception ex)
            {
                _isClosing = false;
                _ = ReportBugAsync("Error during window re-close after task completion", ex);
            }
        }
        catch (Exception ex)
        {
            _isClosing = false;
            _ = ReportBugAsync("Error during window re-close after task completion", ex);
        }
    }

    private const int MaxLogLines = 5000;

    private void LogMessage(string message)
    {
        if (_disposed) return;

        _logChannel.Writer.TryWrite($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
    }

    private async Task ProcessLogsAsync()
    {
        var batch = new List<string>();
        while (await _logChannel.Reader.WaitToReadAsync())
        {
            while (_logChannel.Reader.TryRead(out var log))
            {
                batch.Add(log);
                if (batch.Count >= 50) break; // Process in batches of 50
            }

            if (batch.Count > 0)
            {
                var combinedLogs = string.Join(Environment.NewLine, batch) + Environment.NewLine;
                batch.Clear();

                try
                {
                    if (Application.Current is null) return;

                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        if (_disposed) return;

                        LogViewer.AppendText(combinedLogs);

                        // Clear log if it exceeds the limit to prevent UI freeze from large text operations
                        if (LogViewer.LineCount > MaxLogLines)
                        {
                            var text = LogViewer.Text;
                            var lines = text.Split(Environment.NewLine);
                            if (lines.Length > MaxLogLines)
                            {
                                // Keep only the last 1000 lines instead of clearing everything
                                var keptLines = lines.Skip(lines.Length - 1000).ToArray();
                                LogViewer.Text = string.Join(Environment.NewLine, keptLines) + Environment.NewLine;
                                LogViewer.AppendText($"[{DateTime.Now:HH:mm:ss.fff}] --- Log truncated (kept last 1000 lines) ---{Environment.NewLine}");
                            }
                        }

                        LogViewer.ScrollToEnd();
                    }, System.Windows.Threading.DispatcherPriority.Background);
                }
                catch
                {
                    // Silently fail if the Dispatcher is shutting down
                }
            }

            // Small delay to allow batching more logs if they are arriving rapidly
            await Task.Delay(100);
        }
    }

    private void BrowseInputButton_Click(object sender, RoutedEventArgs e)
    {
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

    private async void StartConversionButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!_dependenciesOk || string.IsNullOrEmpty(_dolphinToolPath))
            {
                LogMessage("Error: Critical dependencies are missing. Cannot start conversion.");
                ShowError("A required file (like DolphinTool.exe) is missing. Please check the application directory.");
                return;
            }

            var inputFolder = InputFolderTextBox.Text;
            var outputFolder = OutputFolderTextBox.Text;
            var deleteFiles = DeleteFilesCheckBox.IsChecked ?? false;
            var useParallelFileProcessing = ParallelProcessingCheckBox.IsChecked == true;

            // Update compression settings from UI
            UpdateBlockSizeFromSelection();

            _currentDegreeOfParallelismForFiles = 1;
            if (useParallelFileProcessing && DegreeOfParallelismComboBox.SelectedItem is System.Windows.Controls.ComboBoxItem selectedItem)
            {
                _ = int.TryParse(selectedItem.Content.ToString(), out _currentDegreeOfParallelismForFiles);
            }

            var inputError = ValidateFolder(inputFolder, "input folder", true);
            if (inputError != null)
            {
                LogMessage($"Error: {inputError}");
                ShowError(inputError);
                return;
            }

            var outputError = ValidateFolder(outputFolder, "output folder", false);
            if (outputError != null)
            {
                LogMessage($"Error: {outputError}");
                ShowError(outputError);
                return;
            }

            if (AreSameFolder(inputFolder, outputFolder))
            {
                const string msg = "The input and output folders must be different directories.";
                LogMessage($"Error: {msg}");
                ShowError(msg);
                return;
            }

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

            ResetOperationStats();
            SetControlsState(false);
            _operationTimer.Restart();

            LogMessage("Starting batch conversion process...");
            LogMessage($"Using DolphinTool: {_dolphinToolPath}");
            LogMessage($"Input folder: {inputFolder}");
            LogMessage($"Output folder: {outputFolder}");
            LogMessage($"Delete original files: {deleteFiles}");
            LogMessage($"Parallel file processing: {useParallelFileProcessing} (Max concurrency: {_currentDegreeOfParallelismForFiles})");
            LogMessage($"RVZ Compression: Method={_rvzCompressionMethod}, Level={_rvzCompressionLevel}, Block Size={_rvzBlockSize}");

            // Wrap the whole job in a task that we can await on exit
            _runningTask = Task.Run(async () =>
            {
                try
                {
                    await PerformBatchConversionAsync(_dolphinToolPath, inputFolder,
                        outputFolder, deleteFiles,
                        useParallelFileProcessing,
                        _currentDegreeOfParallelismForFiles);
                }
                catch (OperationCanceledException)
                {
                    LogMessage("Conversion cancelled by user.");
                }
                catch (Exception ex)
                {
                    LogMessage($"Fatal conversion error: {ex.Message}");
                    await ReportBugAsync("Unhandled exception in conversion", ex);
                }
                finally
                {
                    _operationTimer.Stop();
                    UpdateProcessingTimeDisplay();
                    UpdateWriteSpeedDisplay(0);
                    SetControlsState(true);
                    LogOperationSummary("convert", "Conversion");
                }
            });

            await _runningTask; // keep the UI responsive while running
        }
        catch (Exception ex)
        {
            await ReportBugAsync("Error during StartConversionButton_Click", ex);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        _cts.Cancel();
        LogMessage("Cancellation requested. Waiting for current operation(s) to complete...");

        // Only show the extraction overlay if any extraction is currently in progress
        if (_activeExtractionCount > 0)
        {
            ExtractionOverlayText.Text = "Cancellation requested.\nPlease wait for the current extraction to complete...";
            ExtractionOverlay.Visibility = Visibility.Visible;
        }
    }

    private void SetControlsState(bool enabled)
    {
        // Use Invoke (synchronous) to ensure UI updates complete before returning.
        // This is important when called from background threads (e.g., Task.Run).
        try
        {
            Dispatcher.Invoke(() =>
            {
                MainTabControl.IsEnabled = enabled;

                InputFolderTextBox.IsEnabled = enabled;
                OutputFolderTextBox.IsEnabled = enabled;
                BrowseInputButton.IsEnabled = enabled;
                BrowseOutputButton.IsEnabled = enabled;
                DeleteFilesCheckBox.IsEnabled = enabled;
                ParallelProcessingCheckBox.IsEnabled = enabled;
                StartConversionButton.IsEnabled = enabled;

                VerifyFolderTextBox.IsEnabled = enabled;
                BrowseVerifyFolderButton.IsEnabled = enabled;
                MoveFailedCheckBox.IsEnabled = enabled;
                MoveSuccessCheckBox.IsEnabled = enabled;
                StartVerifyButton.IsEnabled = enabled;

                CancelButton.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;

                if (enabled) // If controls are enabled (operation finished or not started)
                {
                    ClearProgressDisplay(); // Set to idle state

                    // Hide the "Please wait" overlay if it was shown during cancellation
                    ExtractionOverlay.Visibility = Visibility.Collapsed;
                }

                UpdateWriteSpeedDisplay(0);
            });
        }
        catch (TaskCanceledException)
        {
            // Expected during application shutdown
        }
        catch (InvalidOperationException)
        {
            // Dispatcher is shutting down
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

    /// <summary>
    /// Validates a folder path for basic correctness and accessibility.
    /// Returns an error message if validation fails, or null if validation passes.
    /// </summary>
    private static string? ValidateFolder(string folderPath, string label, bool mustExist)
    {
        if (string.IsNullOrWhiteSpace(folderPath))
            return $"Please select the {label}.";

        try
        {
            _ = Path.GetFullPath(folderPath);
        }
        catch (Exception)
        {
            return $"The {label} path is invalid: \"{folderPath}\"";
        }

        switch (mustExist)
        {
            case true when !Directory.Exists(folderPath):
                return $"The {label} does not exist: \"{folderPath}\"";
            case true:
                try
                {
                    _ = Directory.EnumerateFiles(folderPath).Take(1).ToList();
                }
                catch (UnauthorizedAccessException)
                {
                    return $"Access denied to the {label}: \"{folderPath}\"";
                }
                catch (IOException ex)
                {
                    return $"Cannot access the {label}: {ex.Message}";
                }

                break;
        }

        return null;
    }

    /// <summary>
    /// Validates that input and output folders are not the same directory.
    /// </summary>
    private static bool AreSameFolder(string path1, string path2)
    {
        try
        {
            var full1 = Path.GetFullPath(path1).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var full2 = Path.GetFullPath(path2).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return string.Equals(full1, full2, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private async Task PerformBatchConversionAsync(string dolphinToolPath, string inputFolder, string outputFolder, bool deleteFiles, bool useParallelFileProcessing, int maxConcurrency)
    {
        try
        {
            LogMessage("Preparing for batch conversion...");

            var files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(static file => AllSupportedInputExtensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
                .ToArray();

            _totalFilesToProcess = files.Length;
            UpdateStatsDisplay();
            LogMessage($"Found {_totalFilesToProcess} files to process.");
            if (_totalFilesToProcess == 0)
            {
                LogMessage($"No supported files ({SupportedInputExtensionsDisplay}) found in the input folder.");
                return;
            }

            Dispatcher.Invoke(() =>
            {
                ProgressBar.Maximum = Math.Max(_totalFilesToProcess, 1);
                ProgressBar.Value = 0;
            });

            var filesProcessedCount = 0;

            if (useParallelFileProcessing && files.Length > 1)
            {
                // Initialize aggregate speed tracking for parallel processing
                Interlocked.Exchange(ref _aggregateTotalBytesWritten, 0);
                lock (_aggregateSpeedLock)
                {
                    _aggregateLastCheckTime = DateTime.UtcNow;
                }

                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = maxConcurrency,
                    CancellationToken = _cts.Token
                };

                // Start aggregate speed monitoring task
                var speedMonitoringCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token);
                var speedMonitorTask = MonitorAggregateWriteSpeedAsync(speedMonitoringCts.Token);

                await Parallel.ForEachAsync(files, parallelOptions, async (inputFile, token) =>
                {
                    var fileName = Path.GetFileName(inputFile);
                    LogMessage($"[Parallel] Starting: {fileName}");

                    var success = await ProcessFileAsync(dolphinToolPath, inputFile, outputFolder, deleteFiles, true, token);

                    if (success)
                    {
                        Interlocked.Increment(ref _successCount);
                        LogMessage($"[Parallel] Successful: {fileName}");
                    }
                    else
                    {
                        Interlocked.Increment(ref _failureCount);
                        LogMessage($"[Parallel] Failed: {fileName}");
                    }

                    var processed = Interlocked.Increment(ref filesProcessedCount);
                    UpdateProgressDisplay(processed, _totalFilesToProcess, fileName, "Converting");
                    UpdateStatsDisplay();
                    UpdateProcessingTimeDisplay();
                });

                // Stop the aggregate speed monitoring
                speedMonitoringCts.Cancel();
                try
                {
                    await speedMonitorTask;
                }
                catch (OperationCanceledException)
                {
                    // Expected when cancelling the monitoring task
                }
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

                    var success = await ProcessFileAsync(dolphinToolPath, t, outputFolder, deleteFiles, false, _cts.Token);

                    if (success)
                    {
                        Interlocked.Increment(ref _successCount);
                        LogMessage($"Conversion successful: {fileName}");
                    }
                    else
                    {
                        Interlocked.Increment(ref _failureCount);
                        LogMessage($"Conversion failed: {fileName}");
                    }

                    var processed = ++filesProcessedCount;
                    UpdateProgressDisplay(processed, _totalFilesToProcess, fileName, "Converting");
                    UpdateStatsDisplay();
                    UpdateProcessingTimeDisplay();
                }
            }
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

    private async Task<bool> ProcessFileAsync(string dolphinToolPath, string inputFile, string outputFolder, bool deleteOriginal, bool parallelMode, CancellationToken cancellationToken)
    {
        var fileToProcess = inputFile;
        var isArchiveFile = false;
        var tempDir = string.Empty;
        var fileExtension = Path.GetExtension(inputFile);

        try
        {
            if (ArchiveExtensions.Contains(fileExtension, StringComparer.OrdinalIgnoreCase))
            {
                LogMessage($"Processing archive: {Path.GetFileName(inputFile)}");
                var extractResult = await ExtractArchiveAsync(inputFile, cancellationToken);
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
                    // No need to report a bug here, as ExtractArchiveAsync handles bug reporting
                    // for unexpected system errors, while user/archive errors are just logged.
                    return false;
                }
            }

            cancellationToken.ThrowIfCancellationRequested();

            // Get the base name without game image extensions (handles compound extensions like .nkit.iso)
            var outputFile = Path.Combine(outputFolder, GetBaseFileNameWithoutGameExtension(fileToProcess) + ".rvz");

            if (File.Exists(outputFile))
            {
                LogMessage($"Output file already exists, skipping: {Path.GetFileName(outputFile)}");

                // If user requested to delete original files, delete them even when skipping
                if (deleteOriginal)
                {
                    if (isArchiveFile)
                    {
                        // Wait for file handles to be released by OS/antivirus
                        await Task.Delay(2000, CancellationToken.None);
                        await TryDeleteFile(inputFile, $"original archive file: {Path.GetFileName(inputFile)}", CancellationToken.None);
                    }
                    else
                    {
                        // Wait for file handles to be released by OS/antivirus
                        await Task.Delay(2000, CancellationToken.None);
                        await TryDeleteFile(inputFile, $"original ISO file: {Path.GetFileName(inputFile)}", CancellationToken.None);
                    }
                }

                return true;
            }

            var success = await ConvertToRvzAsync(dolphinToolPath, fileToProcess, outputFile, parallelMode, cancellationToken);

            if (!success || !deleteOriginal) return success;

            if (isArchiveFile)
            {
                // Wait for file handles to be released by OS/antivirus
                await Task.Delay(2000, CancellationToken.None);
                await TryDeleteFile(inputFile, $"original archive file: {Path.GetFileName(inputFile)}", CancellationToken.None);
            }
            else
            {
                // Wait for file handles to be released by OS/antivirus
                await Task.Delay(2000, CancellationToken.None);
                await TryDeleteFile(inputFile, $"original ISO file: {Path.GetFileName(inputFile)}", CancellationToken.None);
            }

            return success;
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Processing cancelled for {Path.GetFileName(inputFile)}.");

            var potentialOutputFile = Path.Combine(outputFolder, GetBaseFileNameWithoutGameExtension(fileToProcess) + ".rvz");

            await TryDeleteFile(potentialOutputFile, "partially created RVZ file after cancellation", CancellationToken.None);

            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error processing file {Path.GetFileName(inputFile)}: {ex.Message}");
            await ReportBugAsync($"Error processing file: {Path.GetFileName(inputFile)}", ex);
            var potentialOutputFile = Path.Combine(outputFolder, GetBaseFileNameWithoutGameExtension(fileToProcess) + ".rvz");
            if (File.Exists(potentialOutputFile))
            {
                // Small delay before attempting deletion
                await Task.Delay(100, CancellationToken.None);
                await TryDeleteFile(potentialOutputFile, "partially created RVZ file after error", CancellationToken.None);
            }

            return false;
        }
        finally
        {
            if (isArchiveFile && !string.IsNullOrEmpty(tempDir) && Directory.Exists(tempDir))
            {
                // Small delay before cleaning up temp directory
                await Task.Delay(100, CancellationToken.None);
                TryDeleteDirectory(tempDir, "temporary extraction directory");
            }
        }
    }

    private async Task<bool> ConvertToRvzAsync(string dolphinToolPath, string inputFile, string outputFile, bool parallelMode, CancellationToken cancellationToken)
    {
        // Increment active conversion count for aggregate speed tracking
        Interlocked.Increment(ref _activeConversionCount);

        using var process = new Process();
        try
        {
            LogMessage($"Converting '{Path.GetFileName(inputFile)}' to '{Path.GetFileName(outputFile)}'...");

            var arguments = $"convert -i \"{inputFile}\" -o \"{outputFile}\" -f rvz -c {_rvzCompressionMethod} -l {_rvzCompressionLevel} -b {_rvzBlockSize}";

            process.StartInfo = new ProcessStartInfo
            {
                FileName = dolphinToolPath,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            process.EnableRaisingEvents = true;

            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();

            process.OutputDataReceived += (_, args) =>
            {
                if (string.IsNullOrEmpty(args.Data)) return;

                outputBuilder.AppendLine(args.Data);
                if (!UpdateConversionProgress(args.Data))
                {
                    LogMessage($"[DolphinTool] {args.Data}");
                }
            };
            process.ErrorDataReceived += (_, args) =>
            {
                if (string.IsNullOrEmpty(args.Data))
                {
                    return;
                }

                errorBuilder.AppendLine(args.Data);
                if (!UpdateConversionProgress(args.Data))
                {
                    LogMessage($"[DolphinTool ERROR] {args.Data}");
                }
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            var lastSpeedCheckTime = DateTime.UtcNow;
            long lastFileSize = 0;
            if (File.Exists(outputFile))
            {
                lastFileSize = new FileInfo(outputFile).Length;
            }

            while (!process.HasExited)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        if (!process.HasExited)
                        {
                            // Try graceful termination first
                            process.Kill(true);

                            // Give it a moment to exit gracefully
                            // Use CancellationToken.None to ensure we actually wait for cleanup
                            await Task.Delay(150, CancellationToken.None);

                            // If still running, force kill
                            if (!process.HasExited)
                            {
                                process.Kill();
                                await Task.Delay(100, CancellationToken.None);
                            }
                        }
                    }
                    catch (Exception killEx)
                    {
                        LogMessage($"Error killing process for {Path.GetFileName(inputFile)}: {killEx.Message}");
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                }

                // Wait for either the process to exit or the delay to complete
                var processExitTask = process.WaitForExitAsync(cancellationToken);
                var delayTask = Task.Delay(WriteSpeedUpdateIntervalMs, cancellationToken);
                await Task.WhenAny(processExitTask, delayTask);

                // If process exited, break immediately without waiting for the full delay
                if (process.HasExited || cancellationToken.IsCancellationRequested) break;

                try
                {
                    if (File.Exists(outputFile))
                    {
                        var currentFileSize = new FileInfo(outputFile).Length;
                        var currentTime = DateTime.UtcNow;
                        var timeDelta = currentTime - lastSpeedCheckTime;

                        if (timeDelta.TotalSeconds > 0)
                        {
                            var bytesDelta = currentFileSize - lastFileSize;
                            // Report bytes written to aggregate counter for parallel processing
                            Interlocked.Add(ref _aggregateTotalBytesWritten, bytesDelta);
                            // Only update UI directly if NOT in parallel mode
                            // In parallel mode, MonitorAggregateWriteSpeedAsync handles all UI updates
                            if (!parallelMode)
                            {
                                var speed = (bytesDelta / timeDelta.TotalSeconds) / (1024.0 * 1024.0);
                                UpdateWriteSpeedDisplay(speed);
                            }
                        }

                        lastFileSize = currentFileSize;
                        lastSpeedCheckTime = currentTime;
                    }
                }
                catch (FileNotFoundException)
                {
                    /* File might not be created yet, or deleted */
                }
                catch (Exception ex)
                {
                    LogMessage($"Write speed monitoring error: {ex.Message}");
                }
            }

            // Wait for process exit with cancellation token
            try
            {
                await process.WaitForExitAsync(cancellationToken);
            }
            catch (OperationCanceledException)
            {
                // Handle cancellation during wait
                LogMessage($"Process wait cancelled for {Path.GetFileName(inputFile)}");
                throw;
            }

            LogMessage($"DolphinTool raw output for {Path.GetFileName(inputFile)}: {outputBuilder}");
            if (errorBuilder.Length > 0 && process.ExitCode != 0) LogMessage($"DolphinTool raw error for {Path.GetFileName(inputFile)}: {errorBuilder}");

            return process.ExitCode == 0;
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Conversion cancelled for {Path.GetFileName(inputFile)}.");

            // Wait a moment for the OS to release the file handle after killing the process.
            await Task.Delay(250, CancellationToken.None); // Use CancellationToken.None as the primary token is already cancelled.

            // Wait for the process to fully exit
            if (!process.HasExited)
            {
                try
                {
                    process.Kill(true);
                }
                catch
                {
                    /* ignore */
                }

                // Wait for the process to exit (with timeout)
                try
                {
                    await process.WaitForExitAsync(CancellationToken.None);
                }
                catch
                {
                    /* ignore */
                }
            }

            await TryDeleteFile(outputFile, "partially created RVZ file after cancellation", CancellationToken.None);

            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error converting file {Path.GetFileName(inputFile)}: {ex.Message}");
            await ReportBugAsync($"Error converting file: {Path.GetFileName(inputFile)}", ex);
            if (File.Exists(outputFile))
            {
                // Small delay before attempting deletion
                await Task.Delay(100, CancellationToken.None);
                await TryDeleteFile(outputFile, "partially created RVZ file after error", CancellationToken.None);
            }

            return false;
        }
        finally
        {
            // Decrement active conversion count
            Interlocked.Decrement(ref _activeConversionCount);
            // Only reset display if no more active conversions
            if (Interlocked.CompareExchange(ref _activeConversionCount, 0, 0) == 0)
            {
                UpdateWriteSpeedDisplay(0);
            }
            // Process disposal is handled by 'using' statement
        }
    }

    private async Task<bool> TryDeleteFile(string filePath, string description, CancellationToken cancellationToken)
    {
        try
        {
            if (!File.Exists(filePath)) return true;

            // Clear attributes once up front (handles read-only, hidden, etc.)
            try
            {
                File.SetAttributes(filePath, FileAttributes.Normal);
            }
            catch
            {
                /* Best-effort; delete may still succeed */
            }

            const int maxAttempts = 10;
            for (var attempt = 0; attempt < maxAttempts; attempt++)
            {
                try
                {
                    File.Delete(filePath);
                    LogMessage($"Deleted {description}: {Path.GetFileName(filePath)}");
                    return true;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException && attempt < maxAttempts - 1)
                {
                    // File might be locked or have temporary access issues (e.g. antivirus scanning)

                    // If we've tried several times, try to force a GC to close any potentially abandoned handles
                    // from our own process that might still be hanging onto the file.
                    if (attempt >= 7)
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }

                    if (ex is UnauthorizedAccessException)
                    {
                        // Re-attempt attribute clear in case it changed between retries
                        try
                        {
                            File.SetAttributes(filePath, FileAttributes.Normal);
                        }
                        catch
                        {
                            // ignored
                        }
                    }
                }

                // Exponential backoff: 250ms, 500ms, 750ms, 1000ms, 1500ms, 2000ms, 2500ms, 3000ms, 4000ms, 5000ms
                var delay = attempt switch
                {
                    0 => 250,
                    1 => 500,
                    2 => 750,
                    3 => 1000,
                    4 => 1500,
                    5 => 2000,
                    6 => 2500,
                    7 => 3000,
                    8 => 4000,
                    _ => 5000
                };

                LogMessage($"Cannot delete {Path.GetFileName(filePath)} yet (attempt {attempt + 1}/{maxAttempts}). Waiting {delay}ms...");
                await Task.Delay(delay, cancellationToken);
            }

            LogMessage($"Failed to delete {description} {Path.GetFileName(filePath)} after {maxAttempts} attempts.");
            return false;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error in TryDeleteFile for {description} {filePath}: {ex.Message}");
            _ = Task.Run(() => ReportBugAsync($"Error in TryDeleteFile: {description}", ex), cancellationToken);
            return false;
        }
    }


    private void TryDeleteDirectory(string dirPath, string description)
    {
        try
        {
            if (!Directory.Exists(dirPath)) return;

            for (var i = 0; i < 5; i++)
            {
                try
                {
                    Directory.Delete(dirPath, true);
                    LogMessage($"Cleaned up {description}: {dirPath}");
                    return;
                }
                catch (IOException)
                {
                    Thread.Sleep(50);
                }
            }

            LogMessage($"Failed to clean up {description} {dirPath} after multiple retries.");
        }
        catch (Exception ex)
        {
            LogMessage($"Failed to clean up {description} {dirPath}: {ex.Message}");
        }
    }

    private static string EnsureTrailingSeparator(string path)
    {
        if (path.Length > 0 && path[^1] != Path.DirectorySeparatorChar && path[^1] != Path.AltDirectorySeparatorChar)
            return path + Path.DirectorySeparatorChar;

        return path;
    }

    private static bool IsPathInsideDirectory(string candidatePath, string directoryFullPath)
    {
        var resolved = EnsureTrailingSeparator(Path.GetFullPath(candidatePath));
        return resolved.StartsWith(directoryFullPath, StringComparison.OrdinalIgnoreCase);
    }

    private async Task<(bool Success, string FilePath, string TempDir, string ErrorMessage)> ExtractArchiveAsync(string archivePath, CancellationToken cancellationToken)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        var extension = Path.GetExtension(archivePath);
        var archiveFileName = Path.GetFileName(archivePath);

        // Increment the extraction counter (supports concurrent extractions in parallel mode)
        Interlocked.Increment(ref _activeExtractionCount);

        try
        {
            Directory.CreateDirectory(tempDir);

            if (cancellationToken.IsCancellationRequested)
            {
                LogMessage($"Cancellation requested. Skipping extraction for {archiveFileName}.");
                cancellationToken.ThrowIfCancellationRequested();
            }

            LogMessage($"Extracting {archiveFileName} to temporary directory: {tempDir}");

            if (ArchiveExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            {
                await Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // Normalize temp directory path once for all ZipSlip checks.
                    // Both paths must have trailing separators so StartsWith cannot
                    // match "C:\temp\abc1234" against prefix "C:\temp\abc".
                    var tempDirFullPath = EnsureTrailingSeparator(Path.GetFullPath(tempDir));

                    // Try the standard Archive API first (Seekable)
                    try
                    {
                        using var archive = ArchiveFactory.OpenArchive(archivePath);
                        foreach (var entry in archive.Entries.Where(static e => !e.IsDirectory && !string.IsNullOrEmpty(e.Key)))
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            var entryKey = entry.Key;
                            if (entryKey == null) continue;

                            if (!PrimaryTargetExtensionsInsideArchive.Contains(Path.GetExtension(entryKey), StringComparer.OrdinalIgnoreCase))
                                continue;

                            var entryPath = Path.Combine(tempDir, entryKey);

                            if (!IsPathInsideDirectory(entryPath, tempDirFullPath))
                            {
                                LogMessage($"Skipping potentially malicious archive entry with directory traversal: {entryKey}");
                                continue;
                            }

                            var entryDir = Path.GetDirectoryName(entryPath);
                            if (!string.IsNullOrEmpty(entryDir) && !Directory.Exists(entryDir))
                                Directory.CreateDirectory(entryDir);

                            entry.WriteToFile(entryPath);
                        }
                    }
                    catch (ArchiveException)
                    {
                        LogMessage($"Standard header not found. Trying streaming extraction for {archiveFileName}...");

                        using Stream stream = File.OpenRead(archivePath);
                        using var reader = ReaderFactory.OpenReader(stream, new ReaderOptions());
                        while (reader.MoveToNextEntry())
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            if (reader.Entry.IsDirectory) continue;

                            var entryKey = reader.Entry.Key;
                            if (entryKey == null) continue;

                            if (!PrimaryTargetExtensionsInsideArchive.Contains(Path.GetExtension(entryKey), StringComparer.OrdinalIgnoreCase))
                                continue;

                            var entryPath = Path.Combine(tempDir, entryKey);

                            if (!IsPathInsideDirectory(entryPath, tempDirFullPath))
                            {
                                LogMessage($"Skipping potentially malicious archive entry with directory traversal: {entryKey}");
                                continue;
                            }

                            var entryDir = Path.GetDirectoryName(entryPath);
                            if (!string.IsNullOrEmpty(entryDir) && !Directory.Exists(entryDir))
                                Directory.CreateDirectory(entryDir);

                            using var entryStream = reader.OpenEntryStream();
                            using var fileStream = File.Create(entryPath);
                            entryStream.CopyTo(fileStream);
                        }
                    }

                    if (cancellationToken.IsCancellationRequested)
                        LogMessage($"Extraction of {archiveFileName} completed, but cancellation was requested. Cleaning up.");

                    cancellationToken.ThrowIfCancellationRequested();
                }, cancellationToken);
            }
            else
            {
                // Clean up the temporary directory on unsupported archive type
                TryDeleteDirectory(tempDir, "unsupported archive type extraction directory");
                return (false, string.Empty, string.Empty, $"Unsupported archive type: {extension}");
            }

            // After successful extraction (or if it wasn't cancelled during),
            // look for the target file.
            var supportedFile = Directory.GetFiles(tempDir, "*.*", SearchOption.AllDirectories)
                .FirstOrDefault(static f => PrimaryTargetExtensionsInsideArchive.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

            if (supportedFile != null)
            {
                return (true, supportedFile, tempDir, string.Empty);
            }

            // No supported game image found - clean up the temp directory to prevent disk leak
            TryDeleteDirectory(tempDir, "extraction directory with no supported game images");
            return (false, string.Empty, string.Empty, $"No supported game image ({PrimaryTargetExtensionsDisplay}) found in archive.");
        }
        catch (OperationCanceledException)
        {
            // Log the cancellation
            LogMessage($"Extraction cancelled for {archiveFileName}.");

            // Clean up the temporary directory created for this operation
            TryDeleteDirectory(tempDir, $"cancelled extraction directory for {archiveFileName}");

            // Re-throw to indicate the operation was cancelled
            throw;
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("archive", StringComparison.OrdinalIgnoreCase) ||
                                                   ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase) ||
                                                   ex.Message.Contains("encrypted", StringComparison.OrdinalIgnoreCase) ||
                                                   ex.Message.Contains("determine", StringComparison.OrdinalIgnoreCase) ||
                                                   ex.Message.Contains("stream type", StringComparison.OrdinalIgnoreCase))
        {
            // Handle specific archive errors (corrupt, encrypted, unknown format, etc.) without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. It may be corrupt, encrypted, or an unsupported format. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, string.Empty, errorMessage);
        }
        catch (CryptographicException ex)
        {
            // Handle encrypted archives without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. The archive is encrypted (requires a password), which is not currently supported. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, string.Empty, errorMessage);
        }
        catch (IOException ex) when (ex.Message.Contains("corrupt") || ex.Message.Contains("invalid"))
        {
            // Handle specific archive errors (corrupt, etc.) without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. The archive may be corrupt. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, string.Empty, errorMessage);
        }
        catch (IOException ex) when ((ex.HResult & 0xFFFF) == 0x70 || // ERROR_DISK_FULL (0x80070070)
                                     ex.Message.Contains("not enough space", StringComparison.OrdinalIgnoreCase) ||
                                     ex.Message.Contains("disk full", StringComparison.OrdinalIgnoreCase))
        {
            // Handle disk full errors without sending a bug report - this is a user environment issue.
            var errorMessage = $"Failed to extract archive {archiveFileName}. Not enough disk space available. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, string.Empty, errorMessage);
        }
        catch (SharpCompressException ex)
        {
            // Handle SharpCompress errors (corrupt data, LZMA errors, etc.) without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. The archive data is corrupt or invalid. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, string.Empty, errorMessage);
        }
        catch (Exception ex)
        {
            // Log any other exceptions that occurred during extraction
            LogMessage($"Error extracting archive {archiveFileName}: {ex.Message}");

            // Report the bug asynchronously for unexpected errors
            await ReportBugAsync($"Error extracting archive: {archiveFileName}", ex);

            // Clean up the temporary directory on failure
            TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the error message
            return (false, string.Empty, string.Empty, $"Exception during extraction: {ex.Message}");
        }
        finally
        {
            // Decrement the extraction counter when extraction completes (successfully or with error)
            Interlocked.Decrement(ref _activeExtractionCount);
        }
    }


    private bool UpdateConversionProgress(string progressLine)
    {
        try
        {
            var match = MyRegex().Match(progressLine);
            if (!match.Success) return false;

            var percentageStr = match.Groups[1].Value;
            // FIX: Remove thousand separators and normalize decimal separator for locale-independent parsing
            // Handles formats like "1,000.5%", "1.000,5%", "1000.5%", "1000,5%"
            percentageStr = RemoveThousandSeparators(percentageStr);
            if (!double.TryParse(percentageStr, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                return false;

            // Progress is shown in the progress bar; logging every update causes UI lag
            return true;
        }
        catch (Exception ex)
        {
            LogMessage($"Error parsing DolphinTool progress line '{progressLine}': {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Removes thousand separators and normalizes decimal separator to period for invariant culture parsing.
    /// Handles various locale formats like "1,000.5", "1.000,5", "1000.5", "1000,5".
    /// </summary>
    private static string RemoveThousandSeparators(string numberStr)
    {
        if (string.IsNullOrEmpty(numberStr))
            return numberStr;

        // Count occurrences of comma and period
        var commaCount = numberStr.Count(static c => c == ',');
        var periodCount = numberStr.Count(static c => c == '.');

        // Determine which is the decimal separator based on position
        // The last occurrence of either comma or period is typically the decimal separator
        var lastCommaIndex = numberStr.LastIndexOf(',');
        var lastPeriodIndex = numberStr.LastIndexOf('.');

        switch (commaCount)
        {
            case 0 when periodCount == 0:
                // No separators, return as-is
                return numberStr;
            case > 0 when periodCount == 0:
            {
                // Only commas - determine if decimal or thousand separator
                // If comma is followed by exactly 3 digits at the end, it's likely a thousand separator
                // Otherwise, treat as decimal separator
                if (lastCommaIndex >= 0 && lastCommaIndex < numberStr.Length - 1)
                {
                    var digitsAfterComma = numberStr[(lastCommaIndex + 1)..];
                    if (digitsAfterComma.Length == 3 && digitsAfterComma.All(char.IsDigit))
                    {
                        // Comma is a thousand separator, remove it
                        return numberStr.Replace(",", "");
                    }
                    else
                    {
                        // Comma is a decimal separator, replace with period
                    }
                }

                return numberStr.Replace(',', '.');
            }
        }

        if (periodCount > 0 && commaCount == 0)
        {
            // Only periods - determine if decimal or thousand separator
            if (lastPeriodIndex >= 0 && lastPeriodIndex < numberStr.Length - 1)
            {
                var digitsAfterPeriod = numberStr[(lastPeriodIndex + 1)..];
                if (digitsAfterPeriod.Length == 3 && digitsAfterPeriod.All(char.IsDigit))
                {
                    // Period is a thousand separator, remove it
                    return numberStr.Replace(".", "");
                }
                // Period is already a decimal separator, keep as-is
            }

            return numberStr;
        }

        // Both commas and periods present
        // The rightmost one is the decimal separator
        if (lastCommaIndex > lastPeriodIndex)
        {
            // Comma is the decimal separator
            // Remove periods (thousand separators) and replace comma with period
            return numberStr.Replace(".", "").Replace(',', '.');
        }
        else
        {
            // Period is the decimal separator
            // Remove commas (thousand separators)
            return numberStr.Replace(",", "");
        }
    }

    private void ShowMessageBox(string message, string title, MessageBoxButton buttons, MessageBoxImage icon)
    {
        MessageBox.Show(this, message, title, buttons, icon);
    }

    private void ShowError(string message)
    {
        ShowMessageBox(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private async Task ReportBugAsync(string message, Exception? exception = null)
    {
        try
        {
            var report = new StringBuilder();
            report.AppendLine(message);

            if (exception != null)
            {
                report.AppendLine();
                report.AppendLine("Exception Details:");
                AppendExceptionDetailsToReport(report, exception);
            }

            if (LogViewer != null)
            {
                var logContent = string.Empty;
                await Application.Current.Dispatcher.InvokeAsync(() => logContent = LogViewer.Text);
                if (!string.IsNullOrEmpty(logContent))
                {
                    report.AppendLine().AppendLine("=== Application Log ===").Append(logContent);
                }
            }

            if (App.BugReportServiceInstance != null) await App.BugReportServiceInstance.SendBugReportAsync(report.ToString());
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
            Task.Run(() => ReportBugAsync("Error opening About window", ex));
        }
    }

    private async void CheckForUpdatesMenuItem_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            await CheckForUpdatesAsync(true);
        }
        catch (Exception ex)
        {
            _ = App.BugReportServiceInstance?.SendBugReportAsync($"Error checking for updates: {ex.Message}");
        }
    }

    private async Task CheckForUpdatesAsync(bool isManualCheck)
    {
        LogMessage("Checking for updates...");
        try
        {
            var (isUpdateAvailable, latestRelease) = await _updateService.CheckForUpdatesAsync();

            if (isUpdateAvailable && latestRelease != null)
            {
                LogMessage($"New version available: {latestRelease.Name}");
                var currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                var message = $"A new version ({latestRelease.Name}) is available!\n" +
                              $"You are currently using version {currentVersion}.\n\n" +
                              $"Release Notes:\n{latestRelease.Body}\n\n" +
                              "Would you like to go to the download page?";

                var result = MessageBox.Show(this, message, "Update Available", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    OpenUrl(latestRelease.HtmlUrl);
                }
            }
            else
            {
                LogMessage("You are using the latest version.");
                if (isManualCheck)
                {
                    MessageBox.Show(this, "You are already using the latest version.", "No Updates Found", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }
        catch (Exception ex)
        {
            var errorMessage = $"Error checking for updates: {ex.Message}";
            LogMessage(errorMessage);
            if (isManualCheck)
            {
                MessageBox.Show(this, $"An error occurred while checking for updates:\n{ex.Message}", "Update Check Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            await ReportBugAsync("Failed to check for updates", ex);
        }
    }

    private void OpenUrl(string url)
    {
        try
        {
            var process = Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
            process?.Dispose();
        }
        catch (Exception ex)
        {
            var errorMessage = $"Error opening URL: {url}. Exception: {ex.Message}";
            LogMessage(errorMessage);
            _ = App.BugReportServiceInstance?.SendBugReportAsync(errorMessage);
            MessageBox.Show($"Unable to open link: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ClearProgressDisplay()
    {
        try
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = 0; // Reset progress
                ProgressBar.Maximum = 1; // Ensure the maximum is not zero when idle
                ProgressText.Text = "Ready."; // Set a default idle message
            });
        }
        catch (TaskCanceledException)
        {
            // Expected during application shutdown
        }
        catch (InvalidOperationException)
        {
            // Dispatcher is shutting down
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        _cts.Cancel();
        _cts.Dispose();
        _updateService.Dispose();
        _operationTimer.Stop();
        GC.SuppressFinalize(this);
    }

    private void BrowseVerifyFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var verifyFolder = SelectFolder("Select the folder containing RVZ files to verify");
        if (string.IsNullOrEmpty(verifyFolder)) return;

        VerifyFolderTextBox.Text = verifyFolder;
        LogMessage($"Verification folder selected: {verifyFolder}");
    }

    private async void StartVerifyButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!_dependenciesOk || string.IsNullOrEmpty(_dolphinToolPath))
            {
                LogMessage("Error: Critical dependencies are missing. Cannot start verification.");
                ShowError("A required file (like DolphinTool.exe) is missing. Please check the application directory.");
                return;
            }

            var verifyFolder = VerifyFolderTextBox.Text;
            var useParallelVerification = VerifyParallelProcessingCheckBox.IsChecked == true;

            var maxConcurrency = 1;
            if (useParallelVerification && VerifyDegreeOfParallelismComboBox.SelectedItem is System.Windows.Controls.ComboBoxItem selectedItem)
            {
                _ = int.TryParse(selectedItem.Content.ToString(), out maxConcurrency);
            }

            _moveFailedFiles = MoveFailedCheckBox.IsChecked ?? false;
            _moveSuccessFiles = MoveSuccessCheckBox.IsChecked ?? false;

            var verifyError = ValidateFolder(verifyFolder, "verification folder", true);
            if (verifyError != null)
            {
                LogMessage($"Error: {verifyError}");
                ShowError(verifyError);
                return;
            }

            if (_cts.IsCancellationRequested)
            {
                _cts.Dispose();
                _cts = new CancellationTokenSource();
            }

            ResetOperationStats();
            SetControlsState(false);
            _operationTimer.Restart();

            LogMessage("Starting batch verification process...");
            LogMessage($"Using DolphinTool: {_dolphinToolPath}");
            LogMessage($"Parallel verification: {useParallelVerification} (Max concurrency: {maxConcurrency})");
            LogMessage($"Verification folder: {verifyFolder}");
            if (_moveFailedFiles) LogMessage("Failed files will be moved to '_Failed' subfolder.");
            if (_moveSuccessFiles) LogMessage("Successful files will be moved to '_Success' subfolder.");

            _runningTask = Task.Run(async () =>
            {
                try
                {
                    await PerformBatchVerificationAsync(_dolphinToolPath, verifyFolder,
                        _moveFailedFiles, _moveSuccessFiles, useParallelVerification, maxConcurrency);
                }
                catch (OperationCanceledException)
                {
                    LogMessage("Verification cancelled by user.");
                }
                catch (Exception ex)
                {
                    LogMessage($"Fatal verification error: {ex.Message}");
                    await ReportBugAsync("Unhandled exception in verification", ex);
                }
                finally
                {
                    _operationTimer.Stop();
                    UpdateProcessingTimeDisplay();
                    SetControlsState(true);
                    LogOperationSummary("verify", "Verification");
                }
            });

            await _runningTask;
        }
        catch (Exception ex)
        {
            await ReportBugAsync("Error during StartVerifyButton_Click", ex);
        }
    }

    private async Task PerformBatchVerificationAsync(string dolphinToolPath, string verifyFolder, bool moveFailed, bool moveSuccess, bool useParallel, int maxConcurrency)
    {
        try
        {
            LogMessage("Preparing for batch verification...");

            var files = Directory.GetFiles(verifyFolder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(static file => RvzExtension.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
                .ToArray();

            _totalFilesToProcess = files.Length;
            UpdateStatsDisplay();
            LogMessage($"Found {_totalFilesToProcess} RVZ files to verify.");
            if (_totalFilesToProcess == 0)
            {
                LogMessage($"No RVZ files ({RvzExtensionDisplay}) found in the selected folder.");
                return;
            }

            Dispatcher.Invoke(() =>
            {
                ProgressBar.Maximum = Math.Max(_totalFilesToProcess, 1);
                ProgressBar.Value = 0;
            });

            var filesProcessedCount = 0;

            if (useParallel && files.Length > 1)
            {
                LogMessage("Using parallel verification.");
                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = maxConcurrency,
                    CancellationToken = _cts.Token
                };

                await Parallel.ForEachAsync(files, parallelOptions, async (inputFile, token) =>
                {
                    var success = await VerifyRzvFileAsync(dolphinToolPath, inputFile, verifyFolder, moveFailed, moveSuccess, token);
                    if (success) Interlocked.Increment(ref _successCount);
                    else Interlocked.Increment(ref _failureCount);

                    var processed = Interlocked.Increment(ref filesProcessedCount);
                    UpdateProgressDisplay(processed, _totalFilesToProcess, Path.GetFileName(inputFile), "Verifying");
                    UpdateStatsDisplay();
                    UpdateProcessingTimeDisplay();
                });
            }
            else
            {
                LogMessage("Using sequential verification.");
                foreach (var inputFile in files)
                {
                    if (_cts.Token.IsCancellationRequested) break;

                    var success = await VerifyRzvFileAsync(dolphinToolPath, inputFile, verifyFolder, moveFailed, moveSuccess, _cts.Token);
                    if (success)
                    {
                        Interlocked.Increment(ref _successCount);
                    }
                    else
                    {
                        Interlocked.Increment(ref _failureCount);
                    }

                    var processed = ++filesProcessedCount;
                    UpdateProgressDisplay(processed, _totalFilesToProcess, Path.GetFileName(inputFile), "Verifying");
                    UpdateStatsDisplay();
                    UpdateProcessingTimeDisplay();
                }
            }
        }
        catch (OperationCanceledException)
        {
            LogMessage("Batch verification operation was canceled.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error during batch verification: {ex.Message}");
            ShowError($"Error during batch verification: {ex.Message}");
            await ReportBugAsync("Error during batch verification operation", ex);
        }
    }

    private async Task<bool> VerifyRzvFileAsync(string dolphinToolPath, string inputFile, string baseFolder, bool moveFailed, bool moveSuccess, CancellationToken token)
    {
        var fileName = Path.GetFileName(inputFile);
        using var process = new Process();
        var verificationResult = false;
        string? tempWorkingDirectory = null;
        var wasCanceled = false; // NEW: Flag to track if cancellation occurred for this specific task

        try
        {
            LogMessage($"Verifying: {fileName}...");
            var arguments = $"verify -i \"{inputFile}\"";

            tempWorkingDirectory = Path.Combine(Path.GetTempPath(), "BatchConvertToRVZ_DolphinTool_Temp_" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempWorkingDirectory);

            process.StartInfo = new ProcessStartInfo
            {
                FileName = dolphinToolPath,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = tempWorkingDirectory
            };
            process.EnableRaisingEvents = true;

            var outputBuilder = new StringBuilder();
            process.OutputDataReceived += (_, args) =>
            {
                if (args.Data != null) outputBuilder.AppendLine(args.Data);
            };
            process.ErrorDataReceived += (_, args) =>
            {
                if (args.Data != null) outputBuilder.AppendLine(args.Data);
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            try
            {
                await process.WaitForExitAsync(token);
            }
            catch (OperationCanceledException)
            {
                if (!process.HasExited) process.Kill(true);
                throw;
            }

            var output = outputBuilder.ToString();
            if (process.ExitCode == 0 && output.Contains("Problems Found: No"))
            {
                LogMessage($"[OK] Verification successful for: {fileName}");
                verificationResult = true;
            }
            else
            {
                LogMessage($"[FAIL] Verification failed for: {fileName}. Output: {output.Trim()}");
                verificationResult = false;
            }
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Verification cancelled for {fileName}.");
            wasCanceled = true; // Set the flag
            if (process.HasExited) throw;

            try
            {
                process.Kill(true);
            }
            catch
            {
                /* Ignore */
            }

            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error verifying file {fileName}: {ex.Message}");
            await ReportBugAsync($"Error verifying file: {fileName}", ex);
            verificationResult = false;
        }
        finally
        {
            // NEW: Only move files if the operation was NOT canceled for this specific file
            if (!wasCanceled)
            {
                switch (verificationResult)
                {
                    case true when moveSuccess:
                        await MoveFileToSubfolder(inputFile, baseFolder, "_Success");
                        break;
                    case false when moveFailed:
                        await MoveFileToSubfolder(inputFile, baseFolder, "_Failed");
                        break;
                }
            }
            else
            {
                LogMessage($"Skipping move for '{fileName}' due to cancellation.");
            }

            if (!string.IsNullOrEmpty(tempWorkingDirectory) && Directory.Exists(tempWorkingDirectory))
            {
                TryDeleteDirectory(tempWorkingDirectory, $"temporary working directory for {fileName}");
            }
        }

        return verificationResult;
    }


    private async Task MoveFileToSubfolder(string sourceFilePath, string baseFolder, string subfolderName)
    {
        try
        {
            var destinationFolder = Path.Combine(baseFolder, subfolderName);
            Directory.CreateDirectory(destinationFolder);

            var destinationFilePath = Path.Combine(destinationFolder, Path.GetFileName(sourceFilePath));

            if (File.Exists(destinationFilePath))
            {
                LogMessage($"Deleting existing file at destination: {Path.GetFileName(destinationFilePath)}");
                var deletionSucceeded = await TryDeleteFile(destinationFilePath, $"existing file in {subfolderName} folder", _cts.Token);
                if (!deletionSucceeded)
                {
                    LogMessage($"Cannot move '{Path.GetFileName(sourceFilePath)}' to '{subfolderName}' folder: failed to delete existing file at destination.");
                    return;
                }
            }

            File.Move(sourceFilePath, destinationFilePath);
            LogMessage($"Moved '{Path.GetFileName(sourceFilePath)}' to '{subfolderName}' folder.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error moving file '{Path.GetFileName(sourceFilePath)}' to '{subfolderName}' folder: {ex.Message}");
            _ = Task.Run(() => ReportBugAsync($"Error moving file: {Path.GetFileName(sourceFilePath)} to {subfolderName}", ex));
        }
    }

    private void ResetOperationStats()
    {
        _totalFilesToProcess = 0;
        _successCount = 0;
        _failureCount = 0;
        _operationTimer.Reset();
        // Reset aggregate speed tracking fields
        Interlocked.Exchange(ref _aggregateTotalBytesWritten, 0);
        Interlocked.Exchange(ref _activeConversionCount, 0);
        lock (_aggregateSpeedLock)
        {
            _aggregateLastCheckTime = DateTime.UtcNow;
        }

        UpdateStatsDisplay();
        UpdateProcessingTimeDisplay();
        UpdateWriteSpeedDisplay(0);
        ClearProgressDisplay();
    }

    private void UpdateStatsDisplay()
    {
        try
        {
            Dispatcher.Invoke(() =>
            {
                TotalFilesValue.Text = _totalFilesToProcess.ToString(CultureInfo.InvariantCulture);
                SuccessValue.Text = _successCount.ToString(CultureInfo.InvariantCulture);
                FailedValue.Text = _failureCount.ToString(CultureInfo.InvariantCulture);
            });
        }
        catch (TaskCanceledException)
        {
            // Expected during application shutdown
        }
        catch (InvalidOperationException)
        {
            // Dispatcher is shutting down
        }
    }

    private void UpdateProcessingTimeDisplay()
    {
        var elapsed = _operationTimer.Elapsed;
        try
        {
            Dispatcher.Invoke(() =>
            {
                ProcessingTimeValue.Text = $"{(int)elapsed.TotalHours:D2}:{elapsed:mm\\:ss}";
            });
        }
        catch (TaskCanceledException)
        {
            // Expected during application shutdown
        }
        catch (InvalidOperationException)
        {
            // Dispatcher is shutting down
        }
    }

    private void UpdateWriteSpeedDisplay(double speedInMBps)
    {
        try
        {
            Dispatcher.Invoke(() =>
            {
                WriteSpeedValue.Text = $"{speedInMBps:F1} MB/s";
            });
        }
        catch (TaskCanceledException)
        {
            // Expected during application shutdown
        }
        catch (InvalidOperationException)
        {
            // Dispatcher is shutting down
        }
    }

    /// <summary>
    /// Monitors aggregate write speed across all parallel conversion operations.
    /// This prevents UI flickering by calculating total throughput from all active conversions.
    /// </summary>
    private async Task MonitorAggregateWriteSpeedAsync(CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                await Task.Delay(WriteSpeedUpdateIntervalMs, cancellationToken);

                if (Interlocked.CompareExchange(ref _activeConversionCount, 0, 0) == 0)
                    continue;

                double speedInMBps;
                lock (_aggregateSpeedLock)
                {
                    var currentTime = DateTime.UtcNow;
                    var timeDelta = currentTime - _aggregateLastCheckTime;

                    if (timeDelta.TotalSeconds > 0)
                    {
                        // Interlocked.Exchange atomically reads and resets the counter,
                        // so concurrent Interlocked.Add calls from conversion threads
                        // cannot cause lost updates or torn reads.
                        var bytesThisInterval = Interlocked.Exchange(ref _aggregateTotalBytesWritten, 0);
                        speedInMBps = (bytesThisInterval / timeDelta.TotalSeconds) / (1024.0 * 1024.0);
                        _aggregateLastCheckTime = currentTime;
                    }
                    else
                    {
                        speedInMBps = 0;
                    }
                }

                UpdateWriteSpeedDisplay(Math.Max(0, speedInMBps));
            }
        }
        catch (OperationCanceledException)
        {
            // Expected when the monitoring is cancelled
        }
        catch (Exception ex)
        {
            LogMessage($"Aggregate speed monitoring error: {ex.Message}");
        }
    }

    private void UpdateProgressDisplay(int current, int total, string currentFileName, string operationVerb)
    {
        try
        {
            Dispatcher.Invoke(() =>
            {
                var percentage = total == 0 ? 0 : (double)current / total * 100;
                ProgressText.Text = $"{operationVerb} file {current} of {total}: {currentFileName} ({percentage:F1}%)";
                ProgressBar.Value = current;
            });
        }
        catch (TaskCanceledException)
        {
            // Expected during application shutdown
        }
        catch (InvalidOperationException)
        {
            // Dispatcher is shutting down
        }
    }

    private void LogOperationSummary(string operationVerb, string operationNoun)
    {
        LogMessage("");
        LogMessage($"--- Batch {operationNoun} completed. ---");
        LogMessage($"Total files processed: {_totalFilesToProcess}");
        LogMessage($"Successfully {GetPastTense(operationVerb)}: {_successCount} files");
        if (_failureCount > 0) LogMessage($"Failed to {operationVerb}: {_failureCount} files");

        ShowMessageBox($"Batch {operationNoun} completed.\n\n" +
                       $"Total files processed: {_totalFilesToProcess}\n" +
                       $"Successfully {GetPastTense(operationVerb)}: {_successCount} files\n" +
                       $"Failed: {_failureCount} files",
            $"{operationNoun} Complete", MessageBoxButton.OK,
            _failureCount > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
    }

    private static string GetPastTense(string verb)
    {
        verb = verb.ToLowerInvariant();
        if (verb.EndsWith('y') && verb.Length > 1)
        {
            // Check if the character before 'y' is a consonant (not a vowel)
            var beforeY = verb[^2];
            if (beforeY is not 'a' and not 'e' and not 'i' and not 'o' and not 'u')
            {
                return verb[..^1] + "ied";
            }
        }

        return verb.EndsWith('e') ? verb + "d" : verb + "ed";
    }

    [GeneratedRegex(@"([\d.,]+)%")]
    private static partial Regex MyRegex();

    /// <summary>
    /// Handles compression method selection change.
    /// Updates the compression level slider range based on the selected method.
    /// </summary>
    private void CompressionMethodComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (CompressionMethodComboBox?.SelectedItem is not System.Windows.Controls.ComboBoxItem selectedItem) return;
        if (CompressionLevelSlider == null) return;

        var method = selectedItem.Tag?.ToString() ?? "zstd";
        _rvzCompressionMethod = method;

        // Update compression level range based on selected method
        if (CompressionLevelRanges.TryGetValue(method, out var range))
        {
            CompressionLevelSlider.Minimum = range.Min;
            CompressionLevelSlider.Maximum = range.Max;

            // Adjust current value if it's outside the new range
            if (CompressionLevelSlider.Value < range.Min)
            {
                CompressionLevelSlider.Value = range.Min;
            }
            else if (CompressionLevelSlider.Value > range.Max)
            {
                CompressionLevelSlider.Value = range.Max;
            }
        }

        LogMessage($"Compression method changed to: {method} (level range: {CompressionLevelSlider.Minimum}-{CompressionLevelSlider.Maximum})");
    }

    /// <summary>
    /// Handles compression level slider value change.
    /// Updates the displayed value and stores the setting.
    /// </summary>
    private void CompressionLevelSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (CompressionLevelValue == null) return;

        var level = (int)e.NewValue;
        _rvzCompressionLevel = level;
        CompressionLevelValue.Text = level.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Gets the currently selected block size from the UI.
    /// Called when starting conversion to get the latest value.
    /// </summary>
    private void UpdateBlockSizeFromSelection()
    {
        if (BlockSizeComboBox?.SelectedItem is not System.Windows.Controls.ComboBoxItem selectedItem) return;

        if (selectedItem.Tag != null && int.TryParse(selectedItem.Tag.ToString(), out var blockSize))
        {
            _rvzBlockSize = blockSize;
        }
    }

    /// <summary>
    /// Gets the base file name without game image extensions.
    /// Handles compound extensions like .nkit.iso correctly.
    /// </summary>
    private static string GetBaseFileNameWithoutGameExtension(string filePath)
    {
        var fileName = Path.GetFileName(filePath);

        // Handle compound extension .nkit.iso first
        if (fileName.EndsWith(PrimaryTargetExtensionsInsideArchive[^1], StringComparison.OrdinalIgnoreCase))
        {
            return Path.GetFileNameWithoutExtension(fileName[..^4]); // Remove .iso first, then .nkit
        }

        // Handle other game image extensions
        foreach (var ext in PrimaryTargetExtensionsInsideArchive)
        {
            if (fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
            {
                return fileName[..^ext.Length];
            }
        }

        // Fallback to standard behavior
        return Path.GetFileNameWithoutExtension(fileName);
    }
}