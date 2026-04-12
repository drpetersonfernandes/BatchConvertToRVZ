using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Channels;
using System.Windows;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Collections.Concurrent;
using SharpCompress.Archives;
using SharpCompress.Readers;
using SharpCompress.Common;

namespace BatchConvertToRVZ;

/// <summary>
/// Interaction logic for the MainWindow.
/// Provides batch conversion and verification of game disc images to RVZ format.
/// </summary>
public partial class MainWindow : IDisposable
{
    private bool _disposed;
    private volatile bool _isClosing;
    private volatile bool _isShuttingDown;
    private Task? _runningTask;
    private bool _dependenciesOk;
    private string? _dolphinToolPath;
    private CancellationTokenSource _cts;
    private readonly services.UpdateService _updateService;

    private const string GitHubApiUrl = "https://api.github.com/repos/drpetersonfernandes/BatchConvertToRVZ/releases/latest";

    // Compression settings (now instance variables to allow user configuration)
    private string _rvzCompressionMethod = "zstd"; // Default compression method
    private int _rvzCompressionLevel = 5; // Default compression level
    private int _rvzBlockSize = 131072; // Default block size (128KB)

    // Compression level ranges for different methods
    private static readonly Dictionary<string, (int Min, int Max)> CompressionLevelRanges = new(StringComparer.OrdinalIgnoreCase)
    {
        { "zstd", (1, 22) },
        { "zlib", (1, 9) },
        { "lzma", (1, 9) },
        { "lzma2", (1, 9) },
        { "bzip2", (1, 9) },
        { "lz4", (1, 12) }
    };

    // Supported input extensions (Updated to include all disc images)
    private static readonly string[] AllSupportedInputExtensions = [".iso", ".gcm", ".wbfs", ".rvz", ".zip", ".7z", ".rar"];

    private static readonly string[] ArchiveExtensions = [".zip", ".7z", ".rar"];

    // Primary target extensions inside archives for RVZ conversion
    private static readonly string[] PrimaryTargetExtensionsInsideArchive = [".iso", ".gcm", ".wbfs", ".rvz", ".nkit.iso"];

    // Supported extension for verification
    private static readonly string[] RvzExtension = [".rvz"];

    // Pre-formatted extension lists for user-facing messages
    private static readonly string PrimaryTargetExtensionsDisplay = string.Join(", ", PrimaryTargetExtensionsInsideArchive);

    // Statistics
    private int _totalFilesToProcess;
    private int _successCount;
    private int _failureCount;
    private readonly Stopwatch _operationTimer = new();
    private System.Windows.Threading.DispatcherTimer? _processingTimeUpdateTimer;

    // For Write Speed Calculation
    private const int WriteSpeedUpdateIntervalMs = 500;


    private int _activeConversionCount;

    // Fields for verification move options
    private bool _moveFailedFiles;
    private bool _moveSuccessFiles;

    // Counter to track active extractions (supports concurrent extractions in parallel mode)
    private int _activeExtractionCount;

    // Extraction progress tracking
    private volatile string _currentExtractionFile = string.Empty;
    private long _extractionBytesProcessed;
    private long _extractionTotalBytes;
    private DateTime _extractionLastUpdateTime = DateTime.UtcNow;
    private long _extractionLastBytesProcessed;
    private readonly object _extractionProgressLock = new();

    // Real-time progress tracking for smooth overall progress bar
    private readonly ConcurrentDictionary<string, double> _activeFileProgress = new();
    private readonly object _progressLock = new();

    // File lists for UI
    private readonly BindingList<Models.FileItem> _conversionFiles = new();
    private readonly BindingList<Models.FileItem> _verificationFiles = new();

    private void UpdateOverallProgress()
    {
        try
        {
            if (_totalFilesToProcess == 0) return;

            // Sum up all fractional progress from currently active files
            double activeProgressSum = 0;
            lock (_progressLock)
            {
                foreach (var progress in _activeFileProgress.Values)
                {
                    activeProgressSum += progress / 100.0;
                }
            }

            Dispatcher.BeginInvoke(() =>
            {
                var completed = _successCount + _failureCount;
                // Value is completed files + fraction of currently active ones
                ProgressBar.Value = Math.Min(completed + activeProgressSum, _totalFilesToProcess);
            });
        }
        catch (TaskCanceledException)
        {
        }
        catch (InvalidOperationException)
        {
            // Dispatcher is shutting down
        }
    }

    private void UpdateStatusBar(string status)
    {
        try
        {
            Dispatcher.BeginInvoke(() =>
            {
                if (FindName("StatusBarText") is System.Windows.Controls.TextBlock statusBarText)
                {
                    statusBarText.Text = status;
                }
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


    // Log batching
    private readonly Channel<string> _logChannel = Channel.CreateUnbounded<string>();
    private readonly Task? _logProcessorTask;

    /// <summary>
    /// Initializes a new instance of the <see cref="MainWindow"/> class.
    /// Sets up the UI, logging system, and checks for dependencies.
    /// </summary>
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
        LogMessage("This program will convert GameCube/Wii disc images (.iso, .gcm, .wbfs, .nkit.iso, .rvz) to RVZ format.");
        LogMessage("It also supports extracting game images from ZIP, 7Z, and RAR archives.");
        LogMessage("");
        LogMessage("Use the 'Convert' tab to convert ISO files or archives to RVZ.");
        LogMessage("Use the 'Verify Integrity' tab to check your existing RVZ files.");
        LogMessage("");

        CheckDependencies();

        // Initialize DataGrids
        ConversionFilesDataGrid.ItemsSource = _conversionFiles;
        VerificationFilesDataGrid.ItemsSource = _verificationFiles;

        ResetOperationStats();
        InitializeProcessingTimeTimer();
        Loaded += MainWindow_Loaded;
    }

    private void InitializeProcessingTimeTimer()
    {
        _processingTimeUpdateTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1)
        };
        _processingTimeUpdateTimer.Tick += (_, _) => UpdateProcessingTimeDisplay();
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

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_isClosing)
            return;

        _isShuttingDown = true;
        _logChannel.Writer.TryComplete();

        if (_runningTask is not { IsCompleted: false })
        {
            // No operation running — allow close immediately
            return;
        }

        // An operation is running: cancel it and close once it stops
        e.Cancel = true;
        _isClosing = true;

        _ = Task.Run(async () =>
        {
            try
            {
                await _cts.CancelAsync();

                // Wait up to 5 seconds for the running task to finish
                await Task.WhenAny(_runningTask, Task.Delay(5000));
            }
            catch
            {
                // Ignore all errors during shutdown cancellation
            }
            finally
            {
                // Force-exit on the UI thread regardless of task state
                _ = Dispatcher.BeginInvoke(static () => Application.Current?.Shutdown());
            }
        });
    }

    private const int MaxLogLines = 5000;

    private void LogMessage(string message)
    {
        if (_disposed) return;

        try
        {
            _logChannel.Writer.TryWrite($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
        }
        catch (ChannelClosedException)
        {
            // Channel is closed, ignore
        }
    }

    private async Task ProcessLogsAsync()
    {
        var batch = new List<string>();
        while (!_isShuttingDown && await _logChannel.Reader.WaitToReadAsync())
        {
            while (!_isShuttingDown && _logChannel.Reader.TryRead(out var log))
            {
                batch.Add(log);
                if (batch.Count >= 50) break; // Process in batches of 50
            }

            if (batch.Count > 0)
            {
                if (_isShuttingDown)
                {
                    // Skip UI updates during shutdown to prevent freeze
                    batch.Clear();
                    break;
                }
                else
                {
                    var combinedLogs = string.Join(Environment.NewLine, batch) + Environment.NewLine;
                    batch.Clear();

                    try
                    {
                        if (Application.Current is null) return;

                        await Application.Current.Dispatcher.InvokeAsync(() =>
                        {
                            if (_disposed) return;

                            // Only scroll to end if the user is already at the bottom (or very close to it)
                            // This allows users to scroll up to read previous logs without being snapped back
                            var isAtBottom = LogViewer.VerticalOffset + LogViewer.ViewportHeight >= LogViewer.ExtentHeight - 10;

                            LogViewer.AppendText(combinedLogs);

                            // Efficiently clear log if it exceeds the limit to prevent UI freeze
                            if (LogViewer.LineCount > MaxLogLines)
                            {
                                LogViewer.Clear();
                                LogViewer.AppendText($"[{DateTime.Now:HH:mm:ss.fff}] --- Log cleared (exceeded {MaxLogLines} lines) to prevent UI freeze ---{Environment.NewLine}");
                                isAtBottom = true; // Always scroll to end after clear
                            }

                            if (isAtBottom)
                            {
                                LogViewer.ScrollToEnd();
                            }
                        }, System.Windows.Threading.DispatcherPriority.Background);
                    }
                    catch (Exception ex)
                    {
                        // Silently fail if the Dispatcher is shutting down, but report critical errors
                        if (ex is not InvalidOperationException && ex is not TaskCanceledException)
                        {
                            _ = Task.Run(() => ReportBugAsync("Error in ProcessLogsAsync Dispatcher operation", ex));
                        }
                    }
                }
            }

            // Small delay to allow batching more logs if they are arriving rapidly
            // Check shutdown flag before delay to exit quickly
            if (!_isShuttingDown)
            {
                await Task.Delay(100);
            }
        }
    }

    private void BrowseInputButton_Click(object sender, RoutedEventArgs e)
    {
        var inputFolder = SelectFolder("Select the folder containing ISO files or archives to convert");
        if (string.IsNullOrEmpty(inputFolder)) return;

        InputFolderTextBox.Text = inputFolder;
        LogMessage($"Input folder selected: {inputFolder}");

        PopulateConversionFilesList(inputFolder);
    }

    private void PopulateConversionFilesList(string inputFolder)
    {
        try
        {
            _conversionFiles.Clear();

            var files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(static file => AllSupportedInputExtensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
                .ToArray();

            foreach (var file in files)
            {
                var fileInfo = new FileInfo(file);
                _conversionFiles.Add(new Models.FileItem
                {
                    FileName = Path.GetFileName(file),
                    FullPath = file,
                    FileSize = fileInfo.Length,
                    IsSelected = true
                });
            }

            ConversionFilesDataGrid.ItemsSource = _conversionFiles;
            LogMessage($"Found {_conversionFiles.Count} files in input folder.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error populating file list: {ex.Message}");
        }
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

            // Update compression settings from UI
            UpdateBlockSizeFromSelection();

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
            await SetControlsStateAsync(false);
            _operationTimer.Restart();
            _processingTimeUpdateTimer?.Start();

            LogMessage("Starting batch conversion process...");
            UpdateStatusBar("Starting conversion...");
            LogMessage($"Using DolphinTool: {_dolphinToolPath}");
            LogMessage($"Input folder: {inputFolder}");
            LogMessage($"Output folder: {outputFolder}");
            LogMessage($"Delete original files: {deleteFiles}");
            LogMessage($"RVZ Compression: Method={_rvzCompressionMethod}, Level={_rvzCompressionLevel}, Block Size={_rvzBlockSize}");

            // Get selected files from DataGrid
            var selectedFiles = _conversionFiles.Where(static f => f.IsSelected).Select(static f => f.FullPath).ToArray();
            if (selectedFiles.Length == 0)
            {
                LogMessage("Error: No files selected for conversion.");
                ShowError("Please select at least one file to convert.");
                return;
            }

            // Wrap the whole job in a task that we can await on exit
            try
            {
                _runningTask = Task.Run(async () =>
                {
                    try
                    {
                        await PerformBatchConversionAsync(_dolphinToolPath, selectedFiles,
                            outputFolder, deleteFiles);
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
                });

                await _runningTask.ConfigureAwait(false); // resume on thread pool, not UI thread
            }
            finally
            {
                _operationTimer.Stop();
                _processingTimeUpdateTimer?.Stop();
                UpdateProcessingTimeDisplay();
                UpdateWriteSpeedDisplay(0);
                UpdateStatusBar("Conversion completed");
                await SetControlsStateAsync(true);
                await LogOperationSummaryAsync("convert", "Conversion");
            }
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

    private async Task SetControlsStateAsync(bool enabled)
    {
        // Use InvokeAsync to prevent UI freeze while ensuring UI updates complete.
        // This prevents deadlocks when called from background threads.
        try
        {
            await Dispatcher.InvokeAsync(() =>
            {
                MainTabControl.IsEnabled = enabled;

                InputFolderTextBox.IsEnabled = enabled;
                OutputFolderTextBox.IsEnabled = enabled;
                BrowseInputButton.IsEnabled = enabled;
                BrowseOutputButton.IsEnabled = enabled;
                DeleteFilesCheckBox.IsEnabled = enabled;
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

                    // Update status bar to "Ready"
                    if (FindName("StatusBarText") is System.Windows.Controls.TextBlock statusBarText)
                    {
                        statusBarText.Text = "Ready";
                    }

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

    private async Task PerformBatchConversionAsync(string dolphinToolPath, string[] files, string outputFolder, bool deleteFiles)
    {
        try
        {
            LogMessage("Preparing for batch conversion...");

            _totalFilesToProcess = files.Length;
            UpdateStatsDisplay();
            LogMessage($"Processing {_totalFilesToProcess} selected files.");
            if (_totalFilesToProcess == 0)
            {
                LogMessage("No files selected for conversion.");
                return;
            }

            await Dispatcher.InvokeAsync(() =>
            {
                ProgressBar.Maximum = Math.Max(_totalFilesToProcess, 1);
                ProgressBar.Value = 0;
            });

            var filesProcessedCount = 0;

            LogMessage("Processing files sequentially.");
            foreach (var t in files)
            {
                if (_cts.Token.IsCancellationRequested)
                {
                    LogMessage("Operation canceled by user.");
                    break;
                }

                var fileName = Path.GetFileName(t);

                LogMessage($"Processing: {fileName}");

                var success = await ProcessFileAsync(dolphinToolPath, t, outputFolder, deleteFiles, _cts.Token);

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

                // Remove from active progress AFTER incrementing counters to avoid progress bar jumping backwards
                lock (_progressLock)
                {
                    _activeFileProgress.TryRemove(t, out _);
                }

                UpdateOverallProgress();

                var processed = ++filesProcessedCount;
                UpdateProgressDisplay(processed, _totalFilesToProcess, fileName, "Converting");
                UpdateStatsDisplay();
                UpdateProcessingTimeDisplay();
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

    private async Task<bool> ProcessFileAsync(string dolphinToolPath, string inputFile, string outputFolder, bool deleteOriginal, CancellationToken cancellationToken)
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
                LogMessage($"Output file already exists, overwriting: {Path.GetFileName(outputFile)}");

                // Delete the existing file before creating a new one
                await TryDeleteFile(outputFile, $"existing RVZ file to overwrite: {Path.GetFileName(outputFile)}", cancellationToken);
            }

            var success = await ConvertToRvzAsync(dolphinToolPath, fileToProcess, outputFile, cancellationToken);

            if (!success || !deleteOriginal) return success;

            if (isArchiveFile)
            {
                // Wait for file handles to be released by OS/antivirus
                await Task.Delay(2000, cancellationToken);
                await TryDeleteFile(inputFile, $"original archive file: {Path.GetFileName(inputFile)}", cancellationToken);
            }
            else
            {
                // Wait for file handles to be released by OS/antivirus
                await Task.Delay(2000, cancellationToken);
                await TryDeleteFile(inputFile, $"original ISO file: {Path.GetFileName(inputFile)}", cancellationToken);
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
                await TryDeleteDirectory(tempDir, "temporary extraction directory");
            }
        }
    }

    private async Task<bool> ConvertToRvzAsync(string dolphinToolPath, string inputFile, string outputFile, CancellationToken cancellationToken)
    {
        // Increment active conversion count for aggregate speed tracking
        Interlocked.Increment(ref _activeConversionCount);

        using var process = new Process();

        // Declare event handlers outside try block so they're accessible in finally block
        DataReceivedEventHandler? outputHandler = null;
        DataReceivedEventHandler? errorHandler = null;

        try
        {
            LogMessage($"Converting '{Path.GetFileName(inputFile)}' to '{Path.GetFileName(outputFile)}'...");

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
            process.StartInfo.ArgumentList.Add(_rvzCompressionMethod);
            process.StartInfo.ArgumentList.Add("-l");
            process.StartInfo.ArgumentList.Add(_rvzCompressionLevel.ToString(CultureInfo.InvariantCulture));
            process.StartInfo.ArgumentList.Add("-b");
            process.StartInfo.ArgumentList.Add(_rvzBlockSize.ToString(CultureInfo.InvariantCulture));

            process.EnableRaisingEvents = true;

            // Use ConcurrentQueue for thread-safe output collection from event handlers
            var outputQueue = new ConcurrentQueue<string>();
            var errorQueue = new ConcurrentQueue<string>();

            // Store event handlers so we can remove them later to prevent memory leaks
            outputHandler = (_, args) =>
            {
                if (string.IsNullOrEmpty(args.Data)) return;

                outputQueue.Enqueue(args.Data);
                if (!UpdateConversionProgress(args.Data, inputFile))
                {
                    LogMessage($"[DolphinTool] {args.Data}");
                }
            };
            errorHandler = (_, args) =>
            {
                if (string.IsNullOrEmpty(args.Data))
                {
                    return;
                }

                errorQueue.Enqueue(args.Data);
                if (!UpdateConversionProgress(args.Data, inputFile))
                {
                    LogMessage($"[DolphinTool ERROR] {args.Data}");
                }
            };
            process.OutputDataReceived += outputHandler;
            process.ErrorDataReceived += errorHandler;

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


                            var speed = (bytesDelta / timeDelta.TotalSeconds) / (1024.0 * 1024.0);
                            UpdateWriteSpeedDisplay(speed);
                        }

                        lastFileSize = currentFileSize;
                        lastSpeedCheckTime = currentTime;
                    }
                    else if (lastFileSize > 0)
                    {
                        // File was deleted or moved - reset tracking
                        lastFileSize = 0;
                        lastSpeedCheckTime = DateTime.UtcNow;
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

            // Drain queues to StringBuilder for logging
            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();
            while (outputQueue.TryDequeue(out var line)) outputBuilder.AppendLine(line);
            while (errorQueue.TryDequeue(out var line)) errorBuilder.AppendLine(line);

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
            // Remove event handlers to prevent memory leaks
            process.OutputDataReceived -= outputHandler;
            process.ErrorDataReceived -= errorHandler;

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

            const int maxAttempts = 10;
            for (var attempt = 0; attempt < maxAttempts; attempt++)
            {
                try
                {
                    // Clear attributes on each attempt (handles read-only, hidden, etc.)
                    // This is done inside the loop in case attributes change between retries.
                    try
                    {
                        File.SetAttributes(filePath, FileAttributes.Normal);
                    }
                    catch (Exception attrEx)
                    {
                        // Best-effort; delete might still work even if SetAttributes fails
                        LogMessage($"Warning: Could not clear attributes for {Path.GetFileName(filePath)}: {attrEx.Message}");
                    }

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


    private async Task TryDeleteDirectory(string dirPath, string description)
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
                    await Task.Delay(50);
                }
            }

            LogMessage($"Failed to clean up {description} {dirPath} after multiple retries.");
        }
        catch (Exception ex)
        {
            LogMessage($"Failed to clean up {description} {dirPath}: {ex.Message}");
            await ReportBugAsync($"Error in TryDeleteDirectory: {description}", ex);
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
            UpdateStatusBar($"Extracting {archiveFileName}...");
            LogMessage("Extraction progress will be shown in the status bar...");

            // Set up extraction progress tracking
            lock (_extractionProgressLock)
            {
                _currentExtractionFile = archivePath;
                _extractionBytesProcessed = 0;
                _extractionTotalBytes = 0; // Will be set when we find the entry
                _extractionLastUpdateTime = DateTime.UtcNow;
                _extractionLastBytesProcessed = 0;
            }

            UpdateExtractionProgressDisplay();

            if (ArchiveExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            {
                await Task.Run(async () =>
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

                            // Trim leading slashes and backslashes to ensure Path.Combine works correctly
                            // and doesn't treat the entry as an absolute path (prevents ZipSlip and extraction failures).
                            entryKey = entryKey.TrimStart('/', '\\');

                            if (!HasSupportedGameExtension(entryKey, PrimaryTargetExtensionsInsideArchive))
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

                            // Update total bytes for progress tracking
                            lock (_extractionProgressLock)
                            {
                                _extractionTotalBytes = entry.Size;
                                _extractionBytesProcessed = 0;
                            }

                            UpdateExtractionProgressDisplay();

                            await using var entryStream = await entry.OpenEntryStreamAsync(cancellationToken);
                            await using var fileStream = File.Create(entryPath);
                            await CopyStreamWithCancellationAsync(entryStream, fileStream, cancellationToken, bytesRead =>
                            {
                                lock (_extractionProgressLock)
                                {
                                    _extractionBytesProcessed += bytesRead;
                                }

                                UpdateExtractionProgressDisplay();
                            });
                        }
                    }
                    catch (ArchiveException)
                    {
                        LogMessage($"Standard header not found. Trying streaming extraction for {archiveFileName}...");

                        await using Stream stream = File.OpenRead(archivePath);
                        using var reader = ReaderFactory.OpenReader(stream, new ReaderOptions());
                        while (reader.MoveToNextEntry())
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            if (reader.Entry.IsDirectory) continue;

                            var entryKey = reader.Entry.Key;
                            if (entryKey == null) continue;

                            if (!HasSupportedGameExtension(entryKey, PrimaryTargetExtensionsInsideArchive))
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

                            // Update total bytes for progress tracking
                            lock (_extractionProgressLock)
                            {
                                _extractionTotalBytes = reader.Entry.Size;
                                _extractionBytesProcessed = 0;
                            }

                            UpdateExtractionProgressDisplay();

                            await using var entryStream = reader.OpenEntryStream();
                            await using var fileStream = File.Create(entryPath);
                            await CopyStreamWithCancellationAsync(entryStream, fileStream, cancellationToken, bytesRead =>
                            {
                                lock (_extractionProgressLock)
                                {
                                    _extractionBytesProcessed += bytesRead;
                                }

                                UpdateExtractionProgressDisplay();
                            });
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
                await TryDeleteDirectory(tempDir, "unsupported archive type extraction directory");
                return (false, string.Empty, string.Empty, $"Unsupported archive type: {extension}");
            }

            // After successful extraction (or if it wasn't cancelled during),
            // look for the target file.
            var supportedFile = Directory.GetFiles(tempDir, "*.*", SearchOption.AllDirectories)
                .FirstOrDefault(static f => HasSupportedGameExtension(f, PrimaryTargetExtensionsInsideArchive));

            if (supportedFile != null)
            {
                return (true, supportedFile, tempDir, string.Empty);
            }

            // No supported game image found - clean up the temp directory to prevent disk leak
            await TryDeleteDirectory(tempDir, "extraction directory with no supported game images");
            return (false, string.Empty, string.Empty, $"No supported game image ({PrimaryTargetExtensionsDisplay}) found in archive.");
        }
        catch (OperationCanceledException)
        {
            // Log the cancellation
            LogMessage($"Extraction cancelled for {archiveFileName}.");

            // Clean up the temporary directory created for this operation
            await TryDeleteDirectory(tempDir, $"cancelled extraction directory for {archiveFileName}");

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
            await TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, string.Empty, errorMessage);
        }
        catch (CryptographicException ex)
        {
            // Handle encrypted archives without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. The archive is encrypted (requires a password), which is not currently supported. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            await TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, string.Empty, errorMessage);
        }
        catch (IOException ex) when (ex.Message.Contains("corrupt") || ex.Message.Contains("invalid"))
        {
            // Handle specific archive errors (corrupt, etc.) without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. The archive may be corrupt. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            await TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

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
            await TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, string.Empty, errorMessage);
        }
        catch (SharpCompressException ex)
        {
            // Handle SharpCompress errors (corrupt data, LZMA errors, etc.) without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. The archive data is corrupt or invalid. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            await TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

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
            await TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the error message
            return (false, string.Empty, string.Empty, $"Exception during extraction: {ex.Message}");
        }
        finally
        {
            // Clear extraction progress tracking
            lock (_extractionProgressLock)
            {
                _currentExtractionFile = string.Empty;
                _extractionBytesProcessed = 0;
                _extractionTotalBytes = 0;
                _extractionLastUpdateTime = DateTime.UtcNow;
                _extractionLastBytesProcessed = 0;
            }

            UpdateExtractionProgressDisplay();

            // Update status bar when extraction completes
            UpdateStatusBar("Extraction completed");

            // Decrement the extraction counter when extraction completes (successfully or with error)
            Interlocked.Decrement(ref _activeExtractionCount);
        }
    }

    private static async Task CopyStreamWithCancellationAsync(Stream input, Stream output, CancellationToken cancellationToken, Action<long>? progressCallback = null)
    {
        var buffer = new byte[81920]; // 80KB buffer
        try
        {
            int bytesRead;
            while ((bytesRead = await input.ReadAsync(buffer, cancellationToken).ConfigureAwait(false)) > 0)
            {
                await output.WriteAsync(buffer.AsMemory(0, bytesRead), cancellationToken).ConfigureAwait(false);

                progressCallback?.Invoke(bytesRead);

                cancellationToken.ThrowIfCancellationRequested();
            }

            await output.FlushAsync(cancellationToken);
        }
        catch
        {
            // Ensure output is flushed before re-throwing
            try
            {
                await output.FlushAsync(CancellationToken.None);
            }
            catch
            {
                // ignored
            }

            throw;
        }
    }


    private bool UpdateConversionProgress(string progressLine, string fileName)
    {
        try
        {
            var match = MyRegex().Match(progressLine);
            if (!match.Success) return false;

            var percentageStr = match.Groups[1].Value;
            // FIX: Remove thousand separators and normalize decimal separator for locale-independent parsing
            // Handles formats like "1,000.5%", "1.000,5%", "1000.5%", "1000,5%"
            percentageStr = RemoveThousandSeparators(percentageStr);
            if (!double.TryParse(percentageStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var percentage))
                return false;

            // Store this file's progress and trigger an overall UI update
            lock (_progressLock)
            {
                _activeFileProgress[fileName] = percentage;
            }

            UpdateOverallProgress();

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
    /// Robustly handles various locale formats like "1,000.5", "1.000,5", "1000.5", "1000,5" by treating
    /// the last non-digit character as the decimal separator.
    /// </summary>
    private static string RemoveThousandSeparators(string numberStr)
    {
        if (string.IsNullOrEmpty(numberStr))
            return numberStr;

        // Find the last occurrence of a potential decimal separator (comma or period)
        var lastSeparatorIndex = numberStr.LastIndexOfAny([',', '.']);

        if (lastSeparatorIndex == -1)
        {
            // No separators at all, just return the string (might contain only digits)
            return numberStr;
        }

        // The last separator found is very likely the decimal separator.
        // We take everything before it and remove all non-digits (thousand separators).
        var integerPart = numberStr[..lastSeparatorIndex];
        var decimalPart = numberStr[(lastSeparatorIndex + 1)..];

        // Strip all non-digits from the integer part (removes both dots and commas)
        var sb = new StringBuilder();
        foreach (var c in integerPart)
        {
            if (char.IsDigit(c))
                sb.Append(c);
        }

        // Reconstruct with a period as the decimal separator
        return sb + "." + decimalPart;
    }

    private async Task<MessageBoxResult> ShowMessageBoxAsync(string message, string title, MessageBoxButton buttons, MessageBoxImage icon)
    {
        try
        {
            // Use Dispatcher.InvokeAsync to avoid potential deadlocks when called from async methods
            return await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                // Try to use this window as the owner only if it's still loaded and has a valid handle.
                // This prevents "The calling thread cannot access this object" or "Invalid handle" errors
                // if the window is in the process of closing.
                Window? owner = null;
                try
                {
                    if (!_disposed && IsLoaded)
                    {
                        var helper = new System.Windows.Interop.WindowInteropHelper(this);
                        if (helper.Handle != IntPtr.Zero)
                        {
                            owner = this;
                        }
                    }
                }
                catch
                {
                    // If we can't access the window properties, we'll just show the message box without an owner.
                }

                return owner != null
                    ? MessageBox.Show(owner, message, title, buttons, icon)
                    : MessageBox.Show(message, title, buttons, icon);
            });
        }
        catch (Exception ex)
        {
            LogMessage($"Failed to show message box: {ex.Message}");
            // Fallback: show without owner if something is really wrong
            try
            {
                // Last resort fallback - if we're not on UI thread here, this might still fail in some environments,
                // but we're already in an error state.
                return MessageBox.Show(message, title, buttons, icon);
            }
            catch
            {
                return MessageBoxResult.None;
            }
        }
    }

    // Synchronous wrapper for backward compatibility - uses InvokeAsync internally
    private void ShowMessageBox(string message, string title, MessageBoxButton buttons, MessageBoxImage icon)
    {
        // Use GetAwaiter().GetResult() is safe here because the async operation runs on the UI thread
        // and we're just waiting for it to complete
        ShowMessageBoxAsync(message, title, buttons, icon).GetAwaiter().GetResult();
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

                var result = await ShowMessageBoxAsync(message, "Update Available", MessageBoxButton.YesNo, MessageBoxImage.Information);

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
                    await ShowMessageBoxAsync("You are already using the latest version.", "No Updates Found", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }
        catch (Exception ex)
        {
            var errorMessage = $"Error checking for updates: {ex.Message}";
            LogMessage(errorMessage);
            if (isManualCheck)
            {
                await ShowMessageBoxAsync($"An error occurred while checking for updates:\n{ex.Message}", "Update Check Failed", MessageBoxButton.OK, MessageBoxImage.Error);
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
            ShowMessageBox($"Unable to open link: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ClearProgressDisplay()
    {
        try
        {
            Dispatcher.BeginInvoke(() =>
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

    /// <summary>
    /// Releases all resources used by the <see cref="MainWindow"/>.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        // Complete the log channel to signal the log processor to exit
        _logChannel.Writer.TryComplete();

        // Wait briefly for the log processor to finish (non-blocking)
        if (_logProcessorTask is { IsCompleted: false })
        {
            try
            {
                // Use a short timeout to avoid hanging
                _logProcessorTask.Wait(TimeSpan.FromMilliseconds(500));
            }
            catch
            {
                // Ignore any exceptions during shutdown
            }
        }

        if (!_cts.IsCancellationRequested)
        {
            _cts.Cancel();
        }

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

        PopulateVerificationFilesList(verifyFolder);
    }

    private void PopulateVerificationFilesList(string verifyFolder)
    {
        try
        {
            _verificationFiles.Clear();

            var files = Directory.GetFiles(verifyFolder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(static file => RvzExtension.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
                .ToArray();

            foreach (var file in files)
            {
                var fileInfo = new FileInfo(file);
                _verificationFiles.Add(new Models.FileItem
                {
                    FileName = Path.GetFileName(file),
                    FullPath = file,
                    FileSize = fileInfo.Length,
                    IsSelected = true
                });
            }

            VerificationFilesDataGrid.ItemsSource = _verificationFiles;
            LogMessage($"Found {_verificationFiles.Count} RVZ files in verification folder.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error populating verification file list: {ex.Message}");
        }
    }

    private void SelectAllConversion_Click(object sender, RoutedEventArgs e)
    {
        foreach (var f in _conversionFiles)
        {
            f.IsSelected = true;
        }
    }

    private void DeselectAllConversion_Click(object sender, RoutedEventArgs e)
    {
        foreach (var f in _conversionFiles)
        {
            f.IsSelected = false;
        }
    }

    private void SelectAllVerification_Click(object sender, RoutedEventArgs e)
    {
        foreach (var f in _verificationFiles)
        {
            f.IsSelected = true;
        }
    }

    private void DeselectAllVerification_Click(object sender, RoutedEventArgs e)
    {
        foreach (var f in _verificationFiles)
        {
            f.IsSelected = false;
        }
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
            await SetControlsStateAsync(false);
            _operationTimer.Restart();
            _processingTimeUpdateTimer?.Start();

            // Get selected files from DataGrid
            var selectedFiles = _verificationFiles.Where(static f => f.IsSelected).Select(static f => f.FullPath).ToArray();
            if (selectedFiles.Length == 0)
            {
                LogMessage("Error: No files selected for verification.");
                ShowError("Please select at least one file to verify.");
                return;
            }

            LogMessage("Starting batch verification process...");
            UpdateStatusBar("Starting verification...");
            LogMessage($"Using DolphinTool: {_dolphinToolPath}");

            LogMessage($"Verification folder: {verifyFolder}");
            if (_moveFailedFiles) LogMessage("Failed files will be moved to '_Failed' subfolder.");
            if (_moveSuccessFiles) LogMessage("Successful files will be moved to '_Success' subfolder.");

            _runningTask = Task.Run(async () =>
            {
                try
                {
                    await PerformBatchVerificationAsync(_dolphinToolPath, selectedFiles,
                        _moveFailedFiles, _moveSuccessFiles);
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
            });

            try
            {
                await _runningTask.ConfigureAwait(false); // resume on thread pool, not UI thread
            }
            finally
            {
                _operationTimer.Stop();
                _processingTimeUpdateTimer?.Stop();
                UpdateProcessingTimeDisplay();
                UpdateStatusBar("Verification completed");
                await SetControlsStateAsync(true);
                await LogOperationSummaryAsync("verify", "Verification");
            }
        }
        catch (Exception ex)
        {
            await ReportBugAsync("Error during StartVerifyButton_Click", ex);
        }
    }

    private async Task PerformBatchVerificationAsync(string dolphinToolPath, string[] files, bool moveFailed, bool moveSuccess)
    {
        try
        {
            LogMessage("Preparing for batch verification...");

            _totalFilesToProcess = files.Length;
            UpdateStatsDisplay();
            LogMessage($"Verifying {_totalFilesToProcess} selected RVZ files.");
            if (_totalFilesToProcess == 0)
            {
                LogMessage("No files selected for verification.");
                return;
            }

            await Dispatcher.InvokeAsync(() =>
            {
                ProgressBar.Maximum = Math.Max(_totalFilesToProcess, 1);
                ProgressBar.Value = 0;
            });

            var filesProcessedCount = 0;

            LogMessage("Verifying files sequentially.");
            foreach (var inputFile in files)
            {
                if (_cts.Token.IsCancellationRequested) break;

                var success = await VerifyRzvFileAsync(dolphinToolPath, inputFile, Path.GetDirectoryName(inputFile) ?? string.Empty, moveFailed, moveSuccess, _cts.Token);
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
                UpdateOverallProgress();
                UpdateStatsDisplay();
                UpdateProcessingTimeDisplay();
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

        // Declare event handlers outside try block so they're accessible in finally block
        DataReceivedEventHandler? outputHandler = null;
        DataReceivedEventHandler? errorHandler = null;

        try
        {
            LogMessage($"Verifying: {fileName}...");

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

            // Use ConcurrentQueue for thread-safe output collection from event handlers
            var outputQueue = new ConcurrentQueue<string>();
            outputHandler = (_, args) =>
            {
                if (string.IsNullOrEmpty(args.Data)) return;

                outputQueue.Enqueue(args.Data);
                LogMessage($"[DolphinTool] {args.Data}");
            };
            errorHandler = (_, args) =>
            {
                if (string.IsNullOrEmpty(args.Data)) return;

                outputQueue.Enqueue(args.Data);
                LogMessage($"[DolphinTool ERROR] {args.Data}");
            };
            process.OutputDataReceived += outputHandler;
            process.ErrorDataReceived += errorHandler;

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

            // Drain queue to build output string
            var outputBuilder = new StringBuilder();
            while (outputQueue.TryDequeue(out var line)) outputBuilder.AppendLine(line);
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
                if (process is { HasExited: false })
                {
                    process.Kill(true);
                }
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
            // Remove event handlers to prevent memory leaks
            if (outputHandler != null)
            {
                process.OutputDataReceived -= outputHandler;
            }

            if (errorHandler != null)
            {
                process.ErrorDataReceived -= errorHandler;
            }

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
                await TryDeleteDirectory(tempWorkingDirectory, $"temporary working directory for {fileName}");
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

            // Use try-catch to handle race condition where file might be created between check and move
            try
            {
                File.Move(sourceFilePath, destinationFilePath);
                LogMessage($"Moved '{Path.GetFileName(sourceFilePath)}' to '{subfolderName}' folder.");
            }
            catch (IOException ioEx) when (ioEx.Message.Contains("already exists"))
            {
                // Race condition: file was created between our check and the move
                LogMessage("Destination file appeared during move operation. Attempting to delete and retry...");
                await TryDeleteFile(destinationFilePath, $"race-condition file in {subfolderName} folder", _cts.Token);
                File.Move(sourceFilePath, destinationFilePath);
                LogMessage($"Moved '{Path.GetFileName(sourceFilePath)}' to '{subfolderName}' folder after retry.");
            }
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
        _processingTimeUpdateTimer?.Stop();
        Interlocked.Exchange(ref _activeConversionCount, 0);

        UpdateStatsDisplay();
        UpdateProcessingTimeDisplay();
        UpdateWriteSpeedDisplay(0);
        ClearProgressDisplay();
    }

    private void UpdateStatsDisplay()
    {
        try
        {
            Dispatcher.BeginInvoke(() =>
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
            Dispatcher.BeginInvoke(() =>
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
            Dispatcher.BeginInvoke(() =>
            {
                if (speedInMBps <= 0 && Interlocked.CompareExchange(ref _activeConversionCount, 0, 0) > 0)
                {
                    // Conversion is active but speed is 0 - show "Processing..." instead of "0.0 MB/s"
                    WriteSpeedValue.Text = "Processing...";
                }
                else
                {
                    WriteSpeedValue.Text = $"{speedInMBps:F1} MB/s";
                }
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

    private void UpdateProgressDisplay(int current, int total, string currentFileName, string operationVerb)
    {
        try
        {
            Dispatcher.BeginInvoke(() =>
            {
                var percentage = total == 0 ? 0 : (double)current / total * 100;
                ProgressText.Text = $"{operationVerb} file {current} of {total}: {currentFileName} ({percentage:F1}%)";

                // Also update status bar with simpler message
                if (FindName("StatusBarText") is System.Windows.Controls.TextBlock statusBarText)
                {
                    statusBarText.Text = $"{operationVerb} {currentFileName}...";
                }
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

    private async Task LogOperationSummaryAsync(string operationVerb, string operationNoun)
    {
        LogMessage("");
        LogMessage($"--- Batch {operationNoun} completed. ---");
        LogMessage($"Total files processed: {_totalFilesToProcess}");
        LogMessage($"Successfully {GetPastTense(operationVerb)}: {_successCount} files");
        if (_failureCount > 0) LogMessage($"Failed to {operationVerb}: {_failureCount} files");

        await ShowMessageBoxAsync($"Batch {operationNoun} completed.\n\n" +
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

    private void UpdateExtractionProgressDisplay()
    {
        try
        {
            Dispatcher.BeginInvoke(() =>
            {
                if (!string.IsNullOrEmpty(_currentExtractionFile))
                {
                    var fileName = Path.GetFileName(_currentExtractionFile);
                    var percentage = _extractionTotalBytes > 0
                        ? (double)_extractionBytesProcessed / _extractionTotalBytes * 100
                        : 0;

                    // Calculate extraction speed
                    double speedMBps = 0;
                    lock (_extractionProgressLock)
                    {
                        var currentTime = DateTime.UtcNow;
                        var timeDelta = currentTime - _extractionLastUpdateTime;

                        if (timeDelta.TotalSeconds > 0.5) // Update speed every 0.5 seconds
                        {
                            var bytesDelta = _extractionBytesProcessed - _extractionLastBytesProcessed;
                            speedMBps = (bytesDelta / timeDelta.TotalSeconds) / (1024.0 * 1024.0);

                            _extractionLastUpdateTime = currentTime;
                            _extractionLastBytesProcessed = _extractionBytesProcessed;
                        }
                    }

                    var speedText = speedMBps > 0 ? $" at {speedMBps:F1} MB/s" : "";
                    ProgressText.Text = $"Extracting {fileName}... ({percentage:F1}%){speedText}";

                    // Also update progress bar if we have total bytes
                    if (_extractionTotalBytes > 0)
                    {
                        FileProgressBar.Maximum = 100;
                        FileProgressBar.Value = Math.Min(percentage, 100);
                    }
                }
                else
                {
                    // Reset to ready state if no extraction in progress
                    ProgressText.Text = "Ready.";
                    FileProgressBar.Value = 0;
                    FileProgressBar.Maximum = 1;
                }
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
    /// Checks if a file name has any of the supported extensions, handling compound extensions like .nkit.iso.
    /// </summary>
    private static bool HasSupportedGameExtension(string fileName, string[] supportedExtensions)
    {
        foreach (var ext in supportedExtensions)
        {
            if (fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
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
                _rvzCompressionLevel = range.Min;
            }
            else if (CompressionLevelSlider.Value > range.Max)
            {
                CompressionLevelSlider.Value = range.Max;
                _rvzCompressionLevel = range.Max;
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
            if (blockSize > 0)
            {
                _rvzBlockSize = blockSize;
            }
            else
            {
                LogMessage($"Warning: Invalid block size '{blockSize}' selected. Using default value.");
            }
        }
    }

    /// <summary>
    /// Gets the base file name without game image extensions.
    /// Handles compound extensions like .nkit.iso correctly.
    /// </summary>
    private static string GetBaseFileNameWithoutGameExtension(string filePath)
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