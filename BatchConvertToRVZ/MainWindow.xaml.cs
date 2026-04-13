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
using IDisposable = System.IDisposable;
using StringComparer = System.StringComparer;

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
    private readonly object _ctsLock = new();
    private readonly object _closingLock = new();
    private readonly services.UpdateService _updateService;
    private readonly services.ConversionService _conversionService;
    private readonly services.VerificationService _verificationService;
    private readonly services.ExtractionService _extractionService;
    private readonly services.FileService _fileService;

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

    // Extension arrays moved to FileService

    // Statistics
    private int _totalFilesToProcess;
    private int _successCount;
    private int _failureCount;
    private readonly Stopwatch _operationTimer = new();
    private System.Windows.Threading.DispatcherTimer? _processingTimeUpdateTimer;
    private readonly object _statsLock = new();

    // Write speed calculation
    private long _totalBytesProcessed;
    private DateTime _speedCalculationStartTime;
    private readonly object _speedLock = new();

    // Fields for verification move options
    private bool _moveFailedFiles;
    private bool _moveSuccessFiles;

    // Current operation type for proper cancellation messaging
    private enum OperationType
    {
        None,
        Conversion,
        Verification,
        Extraction
    }

    private OperationType _currentOperation = OperationType.None;

    // File lists for UI
    private readonly BindingList<Models.FileItem> _conversionFiles = new();
    private readonly BindingList<Models.FileItem> _verificationFiles = new();
    private readonly BindingList<Models.FileItem> _extractionFiles = new();

    private void UpdateOverallProgress()
    {
        try
        {
            int totalToProcess;
            int successCount;
            int failureCount;

            lock (_statsLock)
            {
                if (_totalFilesToProcess == 0) return;

                totalToProcess = _totalFilesToProcess;
                successCount = _successCount;
                failureCount = _failureCount;
            }

            Dispatcher.BeginInvoke(() =>
            {
                var completed = successCount + failureCount;
                ProgressBar.Value = Math.Min(completed, totalToProcess);
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

        // Initialize service classes
        _fileService = new services.FileService(LogMessage);
        _conversionService = new services.ConversionService(
            LogMessage,
            message => ReportBugAsync(message),
            _fileService);
        _verificationService = new services.VerificationService(
            LogMessage,
            message => ReportBugAsync(message));
        _extractionService = new services.ExtractionService(
            LogMessage,
            message => ReportBugAsync(message),
            _fileService);

        LogMessage("Welcome to the Batch Convert to RVZ.");
        LogMessage("");
        LogMessage("Use the 'Convert to RVZ' tab to convert ISO, GCM, WBFS, GCZ, WIA or NKIT.ISO files to RVZ.");
        LogMessage("Use the 'Verify Integrity of RVZ' tab to check the integrity of RVZ files.");
        LogMessage("Use the 'Extract from RVZ' tab to convert RVZ files to ISO, WBFS, GCZ or WIA format.");
        LogMessage("");
        LogMessage("");

        CheckDependencies();

        // Initialize DataGrids
        ConversionFilesDataGrid.ItemsSource = _conversionFiles;
        VerificationFilesDataGrid.ItemsSource = _verificationFiles;
        ExtractionFilesDataGrid.ItemsSource = _extractionFiles;

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
            _ = Task.Run(async () =>
            {
                try
                {
                    await ReportBugAsync($"Unsupported platform: {ex.Message}", ex);
                }
                catch
                {
                    /* Silently ignore */
                }
            });
            _dependenciesOk = false;
            StartConversionButton.IsEnabled = false;
            StartVerifyButton.IsEnabled = false;
            return; // Stop further checks
        }

        if (missingFiles.Count != 0)
        {
            _dependenciesOk = false;
            StartConversionButton.IsEnabled = false;
            StartVerifyButton.IsEnabled = false;
            StartExtractionButton.IsEnabled = false;
            var missingFilesString = string.Join(", ", missingFiles);
            var errorMessage = $"The following critical file(s) are missing: {missingFilesString}.\n\nThe application cannot function without them. Please ensure all files from the release archive are in the same folder as this application.";
            LogMessage($"WARNING: {errorMessage.ReplaceLineEndings(" ")}");
            ShowError(errorMessage);
            _ = Task.Run(async () =>
            {
                try
                {
                    await ReportBugAsync($"Missing critical files: {missingFilesString}");
                }
                catch
                {
                    /* Silently ignore */
                }
            });
        }
        else
        {
            _dependenciesOk = true;
            StartConversionButton.IsEnabled = true;
            StartVerifyButton.IsEnabled = true;
            StartExtractionButton.IsEnabled = true;

            if (!string.IsNullOrEmpty(dolphinToolExeName))
            {
                LogMessage($"{dolphinToolExeName} found in the application directory.");
            }

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
        lock (_closingLock)
        {
            if (_isClosing)
            {
                e.Cancel = true;
                return;
            }

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
        }

        _ = Task.Run(async () =>
        {
            try
            {
                lock (_ctsLock)
                {
                    _cts.Cancel();
                }

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
                            _ = Task.Run(async () =>
                            {
                                try
                                {
                                    await ReportBugAsync("Error in ProcessLogsAsync Dispatcher operation", ex);
                                }
                                catch
                                {
                                    /* Silently ignore */
                                }
                            });
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
                .Where(file => _fileService.IsSupportedInputFile(file))
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
            _ = Task.Run(async () =>
            {
                try
                {
                    await ReportBugAsync("Error populating conversion file list", ex);
                }
                catch
                {
                    /* Silently ignore */
                }
            });
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
            // Prevent starting multiple operations simultaneously
            if (_currentOperation != OperationType.None)
            {
                LogMessage($"Error: Cannot start conversion while a {_currentOperation.ToString().ToLowerInvariant()} operation is in progress.");
                ShowError($"Please wait for the current {_currentOperation.ToString().ToLowerInvariant()} operation to complete before starting a new one.");
                return;
            }

            if (!_dependenciesOk || string.IsNullOrEmpty(_dolphinToolPath))
            {
                var exeName = GetDolphinToolExecutableName();
                LogMessage("Error: Critical dependencies are missing. Cannot start conversion.");
                ShowError($"A required file (like {exeName}) is missing. Please check the application directory.");
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

            var selectedFiles = _conversionFiles.Where(static f => f.IsSelected).Select(static f => f.FullPath).ToArray();
            if (selectedFiles.Length == 0)
            {
                LogMessage("Error: No files selected for conversion.");
                ShowError("Please select at least one file to convert.");
                return;
            }

            if (AreSameFolder(inputFolder, outputFolder))
            {
                const string msg = "The input and output folders must be different directories.";
                LogMessage($"Error: {msg}");
                ShowError(msg);
                return;
            }

            if (IsSubdirectory(inputFolder, outputFolder) || IsSubdirectory(outputFolder, inputFolder))
            {
                const string msg = "The input and output folders cannot be nested within each other.";
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

            CancellationToken token;
            lock (_ctsLock)
            {
                // Always refresh the token source for a fresh start using proper disposal pattern
                using (_cts)
                {
                    _cts = new CancellationTokenSource();
                }

                token = _cts.Token;
            }

            // Clear the log before starting the conversion
            await Dispatcher.InvokeAsync(() =>
            {
                LogViewer.Clear();
            });

            ResetOperationStats();
            _currentOperation = OperationType.Conversion;
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

            // Wrap the whole job in a task that we can await on exit
            var wasCancelled = false;
            try
            {
                _runningTask = Task.Run(() => PerformBatchConversionAsync(_dolphinToolPath, selectedFiles, outputFolder, deleteFiles, token), token);

                await _runningTask.ConfigureAwait(false); // resume on thread pool, not UI thread
            }
            catch (OperationCanceledException)
            {
                wasCancelled = true;
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
                _processingTimeUpdateTimer?.Stop();
                UpdateProcessingTimeDisplay();
                UpdateWriteSpeedDisplay(0);
                UpdateStatusBar(wasCancelled ? "Conversion cancelled" : "Conversion completed");
                await SetControlsStateAsync(true);
                if (!wasCancelled)
                {
                    await LogOperationSummaryAsync("convert", "Conversion");
                }
                else
                {
                    LogMessage("--- Batch conversion cancelled. ---");
                }
            }
        }
        catch (Exception ex)
        {
            await ReportBugAsync("Error during StartConversionButton_Click", ex);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        lock (_ctsLock)
        {
            _cts.Cancel();
        }

        LogMessage("Cancellation requested. Waiting for current operation(s) to complete...");

        // Show appropriate overlay text based on current operation type
        var operationName = _currentOperation switch
        {
            OperationType.Conversion => "conversion",
            OperationType.Verification => "verification",
            OperationType.Extraction => "extraction",
            _ => "operation"
        };

        ExtractionOverlayText.Text = $"Cancellation requested.\nPlease wait for the current {operationName} to complete...";
        ExtractionOverlay.Visibility = Visibility.Visible;
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

                ExtractInputFolderTextBox.IsEnabled = enabled;
                ExtractOutputFolderTextBox.IsEnabled = enabled;
                BrowseExtractInputButton.IsEnabled = enabled;
                BrowseExtractOutputButton.IsEnabled = enabled;
                DeleteExtractedFilesCheckBox.IsEnabled = enabled;
                StartExtractionButton.IsEnabled = enabled;

                CancelButton.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;

                if (enabled) // If controls are enabled (operation finished or not started)
                {
                    ClearProgressDisplay(); // Set to idle state
                    _currentOperation = OperationType.None; // Reset operation type

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

    private static bool IsSubdirectory(string parent, string child)
    {
        try
        {
            var parentFull = Path.GetFullPath(parent).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var childFull = Path.GetFullPath(child).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return childFull.StartsWith(parentFull + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                   || childFull.StartsWith(parentFull + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private async Task PerformBatchConversionAsync(string dolphinToolPath, string[] files, string outputFolder, bool deleteFiles, CancellationToken token)
    {
        try
        {
            ResetOperationStats();

            lock (_statsLock)
            {
                _totalFilesToProcess = files.Length;
            }

            int totalFiles;
            lock (_statsLock)
            {
                totalFiles = _totalFilesToProcess;
            }

            await Dispatcher.InvokeAsync(() =>
            {
                FileProgressBar.IsIndeterminate = true;
                FileProgressBar.Value = 0;
                ProgressBar.Maximum = Math.Max(totalFiles, 1);
                ProgressBar.Value = 0;
            });

            lock (_speedLock)
            {
                _speedCalculationStartTime = DateTime.Now;
            }

            // Use the local token variable captured inside the lock for thread safety
            await _conversionService.PerformBatchConversionAsync(
                dolphinToolPath,
                files,
                outputFolder,
                deleteFiles,
                _rvzCompressionMethod,
                _rvzCompressionLevel,
                _rvzBlockSize,
                (processed, total, fileName) =>
                {
                    UpdateProgressDisplay(processed, total, fileName, "Converting");
                    UpdateOverallProgress();
                    UpdateStatsDisplay();
                    UpdateProcessingTimeDisplay();
                    // Track bytes for speed calculation - find file size from conversion files list
                    var fileItem = _conversionFiles.FirstOrDefault(f => f.FileName == fileName);
                    if (fileItem != null)
                    {
                        AddProcessedBytes(fileItem.FileSize);
                    }

                    CalculateAndUpdateWriteSpeed();
                },
                count =>
                {
                    lock (_statsLock)
                    {
                        _successCount += count;
                    }
                },
                count =>
                {
                    lock (_statsLock)
                    {
                        _failureCount += count;
                    }
                },
                token);
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
        finally
        {
            await Dispatcher.InvokeAsync(() =>
            {
                FileProgressBar.IsIndeterminate = false;
                FileProgressBar.Value = 0;
            });
        }
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
        // Check if we're already on the UI thread to avoid deadlock
        if (Application.Current.Dispatcher.CheckAccess())
        {
            // Already on UI thread, show message box directly
            try
            {
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
                    // If we can't access the window properties, we'll just show without an owner.
                }

                if (owner != null)
                    MessageBox.Show(owner, message, title, buttons, icon);
                else
                    MessageBox.Show(message, title, buttons, icon);
            }
            catch
            {
                // Last resort fallback
                MessageBox.Show(message, title, buttons, icon);
            }
        }
        else
        {
            // Not on UI thread, use InvokeAsync
            ShowMessageBoxAsync(message, title, buttons, icon).GetAwaiter().GetResult();
        }
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
            _ = Task.Run(async () =>
            {
                try
                {
                    await ReportBugAsync("Error opening About window", ex);
                }
                catch
                {
                    /* Silently ignore */
                }
            });
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
        catch (System.Net.Http.HttpRequestException ex)
        {
            var errorMessage = $"Failed to check for updates: network error ({ex.Message})";
            LogMessage(errorMessage);
            if (isManualCheck)
            {
                await ShowMessageBoxAsync("Could not connect to update server. Please check your internet connection.", "Update Check Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (TaskCanceledException ex)
        {
            var errorMessage = $"Update check timed out: {ex.Message}";
            LogMessage(errorMessage);
            if (isManualCheck)
            {
                await ShowMessageBoxAsync("Update check timed out. Please try again later.", "Update Check Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
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
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
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
                FileProgressBar.IsIndeterminate = false;
                FileProgressBar.Value = 0;
                FileProgressBar.Maximum = 1;
                ProgressBar.IsIndeterminate = false;
                ProgressBar.Value = 0;
                ProgressBar.Maximum = 1;
                if (FindName("StatusBarText") is System.Windows.Controls.TextBlock statusBarText)
                {
                    statusBarText.Text = "Ready."; // Set a default idle message
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
    /// Releases all resources used by the <see cref="MainWindow"/>.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        // Signal shutdown to ensure log processor exits quickly
        _isShuttingDown = true;

        // Complete the log channel to signal the log processor to exit
        _logChannel.Writer.TryComplete();

        // Wait for the log processor to finish with a reasonable timeout
        if (_logProcessorTask is { IsCompleted: false })
        {
            try
            {
                // Use a longer timeout to ensure all pending logs are processed
                // but don't block indefinitely
                _logProcessorTask.Wait(TimeSpan.FromSeconds(5));
            }
            catch (AggregateException ex) when (ex.InnerException is OperationCanceledException or TaskCanceledException)
            {
                // Expected during shutdown
            }
            catch
            {
                // Ignore any other exceptions during shutdown
            }
        }

        lock (_ctsLock)
        {
            using (_cts)
            {
                if (!_cts.IsCancellationRequested)
                {
                    _cts.Cancel();
                }
            }
        }

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
                .Where(file => _fileService.IsRvzFile(file))
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
            _ = Task.Run(async () =>
            {
                try
                {
                    await ReportBugAsync("Error populating verification file list", ex);
                }
                catch
                {
                    /* Silently ignore */
                }
            });
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
            // Prevent starting multiple operations simultaneously
            if (_currentOperation != OperationType.None)
            {
                LogMessage($"Error: Cannot start verification while a {_currentOperation.ToString().ToLowerInvariant()} operation is in progress.");
                ShowError($"Please wait for the current {_currentOperation.ToString().ToLowerInvariant()} operation to complete before starting a new one.");
                return;
            }

            if (!_dependenciesOk || string.IsNullOrEmpty(_dolphinToolPath))
            {
                var exeName = GetDolphinToolExecutableName();
                LogMessage("Error: Critical dependencies are missing. Cannot start verification.");
                ShowError($"A required file (like {exeName}) is missing. Please check the application directory.");
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

            var selectedFiles = _verificationFiles.Where(static f => f.IsSelected).Select(static f => f.FullPath).ToArray();
            if (selectedFiles.Length == 0)
            {
                LogMessage("Error: No files selected for verification.");
                ShowError("Please select at least one file to verify.");
                return;
            }

            CancellationToken token;
            lock (_ctsLock)
            {
                // Always refresh the token source for a fresh start using proper disposal pattern
                using (_cts)
                {
                    _cts = new CancellationTokenSource();
                }

                token = _cts.Token;
            }

            // Clear the log before starting the verification
            await Dispatcher.InvokeAsync(() =>
            {
                LogViewer.Clear();
            });

            ResetOperationStats();
            _currentOperation = OperationType.Verification;
            await SetControlsStateAsync(false);
            _operationTimer.Restart();
            _processingTimeUpdateTimer?.Start();

            LogMessage("Starting batch verification process...");
            UpdateStatusBar("Starting verification...");
            LogMessage($"Using DolphinTool: {_dolphinToolPath}");

            LogMessage($"Verification folder: {verifyFolder}");
            if (_moveFailedFiles) LogMessage("Failed files will be moved to '_Failed' subfolder.");
            if (_moveSuccessFiles) LogMessage("Successful files will be moved to '_Success' subfolder.");

            var wasCancelled = false;
            _runningTask = Task.Run(() => PerformBatchVerificationAsync(_dolphinToolPath, selectedFiles, _moveFailedFiles, _moveSuccessFiles, token), token);

            try
            {
                await _runningTask.ConfigureAwait(false); // resume on thread pool, not UI thread
            }
            catch (OperationCanceledException)
            {
                wasCancelled = true;
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
                _processingTimeUpdateTimer?.Stop();
                UpdateProcessingTimeDisplay();
                UpdateStatusBar(wasCancelled ? "Verification cancelled" : "Verification completed");
                await SetControlsStateAsync(true);
                if (!wasCancelled)
                {
                    await LogOperationSummaryAsync("verify", "Verification");
                }
                else
                {
                    LogMessage("--- Batch verification cancelled. ---");
                }

                // Refresh the verification file list to reflect any moved files
                if (!string.IsNullOrEmpty(VerifyFolderTextBox.Text) && Directory.Exists(VerifyFolderTextBox.Text))
                {
                    PopulateVerificationFilesList(VerifyFolderTextBox.Text);
                }
            }
        }
        catch (Exception ex)
        {
            await ReportBugAsync("Error during StartVerifyButton_Click", ex);
        }
    }

    private async Task PerformBatchVerificationAsync(string dolphinToolPath, string[] files, bool moveFailed, bool moveSuccess, CancellationToken token)
    {
        try
        {
            ResetOperationStats();

            lock (_statsLock)
            {
                _totalFilesToProcess = files.Length;
            }

            int totalFiles;
            lock (_statsLock)
            {
                totalFiles = _totalFilesToProcess;
            }

            await Dispatcher.InvokeAsync(() =>
            {
                FileProgressBar.IsIndeterminate = true;
                FileProgressBar.Value = 0;
                ProgressBar.Maximum = Math.Max(totalFiles, 1);
                ProgressBar.Value = 0;
            });

            lock (_speedLock)
            {
                _speedCalculationStartTime = DateTime.Now;
            }

            // Use the local token variable captured inside the lock for thread safety
            await _verificationService.PerformBatchVerificationAsync(
                dolphinToolPath,
                files,
                moveFailed,
                moveSuccess,
                (processed, total, fileName) =>
                {
                    UpdateProgressDisplay(processed, total, fileName, "Verifying");
                    UpdateOverallProgress();
                    UpdateStatsDisplay();
                    UpdateProcessingTimeDisplay();
                    // Track bytes for speed calculation - find file size from verification files list
                    var fileItem = _verificationFiles.FirstOrDefault(f => f.FileName == fileName);
                    if (fileItem != null)
                    {
                        AddProcessedBytes(fileItem.FileSize);
                    }

                    CalculateAndUpdateWriteSpeed();
                },
                count =>
                {
                    lock (_statsLock)
                    {
                        _successCount += count;
                    }
                },
                count =>
                {
                    lock (_statsLock)
                    {
                        _failureCount += count;
                    }
                },
                token);
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
        finally
        {
            await Dispatcher.InvokeAsync(() =>
            {
                FileProgressBar.IsIndeterminate = false;
                FileProgressBar.Value = 0;
            });
        }
    }

    private void ResetOperationStats()
    {
        lock (_statsLock)
        {
            _totalFilesToProcess = 0;
            _successCount = 0;
            _failureCount = 0;
        }

        _operationTimer.Reset();
        _processingTimeUpdateTimer?.Stop();

        // Reset speed calculation
        lock (_speedLock)
        {
            _totalBytesProcessed = 0;
            _speedCalculationStartTime = DateTime.Now;
        }

        UpdateStatsDisplay();
        UpdateProcessingTimeDisplay();
        UpdateWriteSpeedDisplay(0);
        ClearProgressDisplay();
    }

    /// <summary>
    /// Adds processed bytes to the speed calculation.
    /// </summary>
    /// <param name="bytes">Number of bytes processed.</param>
    private void AddProcessedBytes(long bytes)
    {
        lock (_speedLock)
        {
            _totalBytesProcessed += bytes;
        }
    }

    /// <summary>
    /// Calculates and updates the current write speed display.
    /// </summary>
    private void CalculateAndUpdateWriteSpeed()
    {
        double speedInMBps;
        lock (_speedLock)
        {
            var elapsed = DateTime.Now - _speedCalculationStartTime;
            if (elapsed.TotalSeconds > 0 && _totalBytesProcessed > 0)
            {
                speedInMBps = _totalBytesProcessed / (elapsed.TotalSeconds * 1024 * 1024);
            }
            else
            {
                speedInMBps = 0;
            }
        }

        UpdateWriteSpeedDisplay(speedInMBps);
    }

    private void UpdateStatsDisplay()
    {
        try
        {
            int totalFiles;
            int successCount;
            int failureCount;

            lock (_statsLock)
            {
                totalFiles = _totalFilesToProcess;
                successCount = _successCount;
                failureCount = _failureCount;
            }

            Dispatcher.BeginInvoke(() =>
            {
                TotalFilesValue.Text = totalFiles.ToString(CultureInfo.InvariantCulture);
                SuccessValue.Text = successCount.ToString(CultureInfo.InvariantCulture);
                FailedValue.Text = failureCount.ToString(CultureInfo.InvariantCulture);
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

    private void UpdateProgressDisplay(int current, int total, string currentFileName, string operationVerb)
    {
        try
        {
            Dispatcher.BeginInvoke(() =>
            {
                var percentage = total == 0 ? 0 : (double)current / total * 100;
                if (FindName("StatusBarText") is System.Windows.Controls.TextBlock statusBarText)
                {
                    statusBarText.Text = $"{operationVerb} file {current} of {total}: {currentFileName} ({percentage:F1}%)";
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
        int totalFiles;
        int successCount;
        int failureCount;

        lock (_statsLock)
        {
            totalFiles = _totalFilesToProcess;
            successCount = _successCount;
            failureCount = _failureCount;
        }

        LogMessage("");
        LogMessage($"--- Batch {operationNoun} completed. ---");
        LogMessage($"Total files processed: {totalFiles}");
        LogMessage($"Successfully {GetPastTense(operationVerb)}: {successCount} files");
        if (failureCount > 0) LogMessage($"Failed to {operationVerb}: {failureCount} files");

        await ShowMessageBoxAsync($"Batch {operationNoun} completed.\n\n" +
                                  $"Total files processed: {totalFiles}\n" +
                                  $"Successfully {GetPastTense(operationVerb)}: {successCount} files\n" +
                                  $"Failed: {failureCount} files",
            $"{operationNoun} Complete", MessageBoxButton.OK,
            failureCount > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
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

        // Validate level is within allowed range for the selected compression method
        if (CompressionMethodComboBox?.SelectedItem is System.Windows.Controls.ComboBoxItem selectedItem)
        {
            var method = selectedItem.Tag?.ToString() ?? "zstd";

            if (CompressionLevelRanges.TryGetValue(method, out var range))
            {
                // Ensure level is within valid range
                if (level < range.Min)
                {
                    level = range.Min;
                    CompressionLevelSlider.Value = level;
                }
                else if (level > range.Max)
                {
                    level = range.Max;
                    CompressionLevelSlider.Value = level;
                }
            }
        }

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
            // Add upper bound check (2MB) based on the defined values in the combo box
            if (blockSize is > 0 and <= 2097152)
            {
                _rvzBlockSize = blockSize;
            }
            else
            {
                LogMessage($"Warning: Invalid block size '{blockSize}' selected. Using default value.");
                _rvzBlockSize = 131072; // Default to 128KB
            }
        }
    }

    #region Extraction Tab Event Handlers

    private void BrowseExtractInputButton_Click(object sender, RoutedEventArgs e)
    {
        var inputFolder = SelectFolder("Select the folder containing RVZ files to extract");
        if (string.IsNullOrEmpty(inputFolder)) return;

        ExtractInputFolderTextBox.Text = inputFolder;
        LogMessage($"Extraction input folder selected: {inputFolder}");

        PopulateExtractionFilesList(inputFolder);
    }

    private void PopulateExtractionFilesList(string inputFolder)
    {
        try
        {
            _extractionFiles.Clear();

            var files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(file => _fileService.IsRvzFile(file))
                .ToArray();

            foreach (var file in files)
            {
                var fileInfo = new FileInfo(file);
                _extractionFiles.Add(new Models.FileItem
                {
                    FileName = Path.GetFileName(file),
                    FullPath = file,
                    FileSize = fileInfo.Length,
                    IsSelected = true
                });
            }

            ExtractionFilesDataGrid.ItemsSource = _extractionFiles;
            LogMessage($"Found {_extractionFiles.Count} RVZ files in extraction folder.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error populating extraction file list: {ex.Message}");
            _ = Task.Run(async () =>
            {
                try
                {
                    await ReportBugAsync("Error populating extraction file list", ex);
                }
                catch
                {
                    /* Silently ignore */
                }
            });
        }
    }

    private void BrowseExtractOutputButton_Click(object sender, RoutedEventArgs e)
    {
        var outputFolder = SelectFolder("Select the output folder where ISO files will be saved");
        if (string.IsNullOrEmpty(outputFolder)) return;

        ExtractOutputFolderTextBox.Text = outputFolder;
        LogMessage($"Extraction output folder selected: {outputFolder}");
    }

    private void SelectAllExtraction_Click(object sender, RoutedEventArgs e)
    {
        foreach (var f in _extractionFiles)
        {
            f.IsSelected = true;
        }
    }

    private void DeselectAllExtraction_Click(object sender, RoutedEventArgs e)
    {
        foreach (var f in _extractionFiles)
        {
            f.IsSelected = false;
        }
    }

    private async void StartExtractionButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Prevent starting multiple operations simultaneously
            if (_currentOperation != OperationType.None)
            {
                LogMessage($"Error: Cannot start extraction while a {_currentOperation.ToString().ToLowerInvariant()} operation is in progress.");
                ShowError($"Please wait for the current {_currentOperation.ToString().ToLowerInvariant()} operation to complete before starting a new one.");
                return;
            }

            if (!_dependenciesOk || string.IsNullOrEmpty(_dolphinToolPath))
            {
                var exeName = GetDolphinToolExecutableName();
                LogMessage("Error: Critical dependencies are missing. Cannot start extraction.");
                ShowError($"A required file (like {exeName}) is missing. Please check the application directory.");
                return;
            }

            var inputFolder = ExtractInputFolderTextBox.Text;
            var outputFolder = ExtractOutputFolderTextBox.Text;
            var deleteFiles = DeleteExtractedFilesCheckBox.IsChecked ?? false;
            var outputFormat = ExtractOutputFormatComboBox.Text.ToLowerInvariant();

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

            var selectedFiles = _extractionFiles.Where(static f => f.IsSelected).Select(static f => f.FullPath).ToArray();
            if (selectedFiles.Length == 0)
            {
                LogMessage("Error: No files selected for extraction.");
                ShowError("Please select at least one file to extract.");
                return;
            }

            if (AreSameFolder(inputFolder, outputFolder))
            {
                const string msg = "The input and output folders must be different directories.";
                LogMessage($"Error: {msg}");
                ShowError(msg);
                return;
            }

            if (IsSubdirectory(inputFolder, outputFolder) || IsSubdirectory(outputFolder, inputFolder))
            {
                const string msg = "The input and output folders cannot be nested within each other.";
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

            CancellationToken token;
            lock (_ctsLock)
            {
                // Always refresh the token source for a fresh start using proper disposal pattern
                using (_cts)
                {
                    _cts = new CancellationTokenSource();
                }

                token = _cts.Token;
            }

            // Clear the log before starting the extraction
            await Dispatcher.InvokeAsync(() =>
            {
                LogViewer.Clear();
            });

            ResetOperationStats();
            _currentOperation = OperationType.Extraction;
            await SetControlsStateAsync(false);
            _operationTimer.Restart();
            _processingTimeUpdateTimer?.Start();

            LogMessage("Starting batch extraction process...");
            UpdateStatusBar("Starting extraction...");
            LogMessage($"Using DolphinTool: {_dolphinToolPath}");
            LogMessage($"Input folder: {inputFolder}");
            LogMessage($"Output folder: {outputFolder}");
            LogMessage($"Output format: {outputFormat.ToUpperInvariant()}");
            LogMessage($"Delete original files: {deleteFiles}");

            var wasCancelled = false;
            try
            {
                _runningTask = Task.Run(() => PerformBatchExtractionAsync(_dolphinToolPath, selectedFiles, outputFolder, deleteFiles, outputFormat, token), token);

                await _runningTask.ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                wasCancelled = true;
                LogMessage("Extraction cancelled by user.");
            }
            catch (Exception ex)
            {
                LogMessage($"Fatal extraction error: {ex.Message}");
                await ReportBugAsync("Unhandled exception in extraction", ex);
            }
            finally
            {
                _operationTimer.Stop();
                _processingTimeUpdateTimer?.Stop();
                UpdateProcessingTimeDisplay();
                UpdateWriteSpeedDisplay(0);
                UpdateStatusBar(wasCancelled ? "Extraction cancelled" : "Extraction completed");
                await SetControlsStateAsync(true);
                if (!wasCancelled)
                {
                    await LogOperationSummaryAsync("extract", "Extraction");
                }
                else
                {
                    LogMessage("--- Batch extraction cancelled. ---");
                }
            }
        }
        catch (Exception ex)
        {
            await ReportBugAsync("Error during StartExtractionButton_Click", ex);
        }
    }

    private async Task PerformBatchExtractionAsync(string dolphinToolPath, string[] files, string outputFolder, bool deleteFiles, string outputFormat, CancellationToken token)
    {
        try
        {
            ResetOperationStats();

            lock (_statsLock)
            {
                _totalFilesToProcess = files.Length;
            }

            int totalFiles;
            lock (_statsLock)
            {
                totalFiles = _totalFilesToProcess;
            }

            await Dispatcher.InvokeAsync(() =>
            {
                FileProgressBar.IsIndeterminate = true;
                FileProgressBar.Value = 0;
                ProgressBar.Maximum = Math.Max(totalFiles, 1);
                ProgressBar.Value = 0;
            });

            lock (_speedLock)
            {
                _speedCalculationStartTime = DateTime.Now;
            }

            // Use the local token variable captured inside the lock for thread safety
            await _extractionService.PerformBatchExtractionAsync(
                dolphinToolPath,
                files,
                outputFolder,
                deleteFiles,
                outputFormat,
                (processed, total, fileName) =>
                {
                    UpdateProgressDisplay(processed, total, fileName, "Extracting");
                    UpdateOverallProgress();
                    UpdateStatsDisplay();
                    UpdateProcessingTimeDisplay();
                    // Track bytes for speed calculation - find file size from extraction files list
                    var fileItem = _extractionFiles.FirstOrDefault(f => f.FileName == fileName);
                    if (fileItem != null)
                    {
                        AddProcessedBytes(fileItem.FileSize);
                    }

                    CalculateAndUpdateWriteSpeed();
                },
                count =>
                {
                    lock (_statsLock)
                    {
                        _successCount += count;
                    }
                },
                count =>
                {
                    lock (_statsLock)
                    {
                        _failureCount += count;
                    }
                },
                token);
        }
        catch (OperationCanceledException)
        {
            LogMessage("Batch extraction operation was canceled.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error during batch extraction: {ex.Message}");
            ShowError($"Error during batch extraction: {ex.Message}");
            await ReportBugAsync("Error during batch extraction operation", ex);
        }
        finally
        {
            await Dispatcher.InvokeAsync(() =>
            {
                FileProgressBar.IsIndeterminate = false;
                FileProgressBar.Value = 0;
            });
        }
    }

    #endregion

    #region Drag and Drop Event Handlers

    private void ConversionFilesDataGrid_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private void ConversionFilesDataGrid_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
            {
                HandleDroppedFiles(files, "conversion");
            }
        }
    }

    private void VerificationFilesDataGrid_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private void VerificationFilesDataGrid_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
            {
                HandleDroppedFiles(files, "verification");
            }
        }
    }

    private void ExtractionFilesDataGrid_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private void ExtractionFilesDataGrid_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
            {
                HandleDroppedFiles(files, "extraction");
            }
        }
    }

    private void HandleDroppedFiles(string[] files, string target)
    {
        if (files.Length == 0) return;

        try
        {
            var folderPath = files.FirstOrDefault(Directory.Exists);

            if (folderPath != null)
            {
                // If a folder was dropped, use it as the input folder
                switch (target)
                {
                    case "conversion":
                        InputFolderTextBox.Text = folderPath;
                        LogMessage($"Input folder selected via drag-and-drop: {folderPath}");
                        PopulateConversionFilesList(folderPath);
                        break;
                    case "verification":
                        VerifyFolderTextBox.Text = folderPath;
                        LogMessage($"Verification folder selected via drag-and-drop: {folderPath}");
                        PopulateVerificationFilesList(folderPath);
                        break;
                    case "extraction":
                        ExtractInputFolderTextBox.Text = folderPath;
                        LogMessage($"Extraction input folder selected via drag-and-drop: {folderPath}");
                        PopulateExtractionFilesList(folderPath);
                        break;
                }
            }
            else
            {
                // Individual files dropped - add them to the appropriate list
                var fileList = files.Where(File.Exists).ToList();
                if (fileList.Count > 0)
                {
                    LogMessage($"Dropped {fileList.Count} file(s) for {target}.");
                    AddIndividualFilesToList(fileList, target);
                }
            }
        }
        catch (Exception ex)
        {
            LogMessage($"Error handling dropped files: {ex.Message}");
            ShowError($"Error processing dropped files: {ex.Message}");
        }
    }

    /// <summary>
    /// Adds individual files to the appropriate file list based on the target.
    /// </summary>
    /// <param name="files">The list of file paths to add.</param>
    /// <param name="target">The target list ("conversion", "verification", or "extraction").</param>
    private void AddIndividualFilesToList(List<string> files, string target)
    {
        try
        {
            // Filter files by supported extensions based on target
            var filteredFiles = target switch
            {
                "conversion" => files.Where(file => _fileService.IsSupportedInputFile(file)),
                "verification" => files.Where(file => _fileService.IsRvzFile(file)),
                "extraction" => files.Where(file => _fileService.IsSupportedExtractionInputFile(file)),
                _ => files
            };

            var fileArray = filteredFiles.ToArray();
            if (fileArray.Length == 0)
            {
                ShowMessageBox("No supported files were found in the drop. Please check file extensions.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Get the appropriate list and determine the common parent directory
            var fileList = target switch
            {
                "conversion" => _conversionFiles,
                "verification" => _verificationFiles,
                "extraction" => _extractionFiles,
                _ => null
            };

            if (fileList == null) return;

            // Find common parent directory for the text box display
            string? commonDirectory = null;
            if (fileArray.Length > 0)
            {
                commonDirectory = Path.GetDirectoryName(fileArray[0]);
                for (var i = 1; i < fileArray.Length && !string.IsNullOrEmpty(commonDirectory); i++)
                {
                    var dir = Path.GetDirectoryName(fileArray[i]);
                    while (!string.IsNullOrEmpty(dir) && !string.IsNullOrEmpty(commonDirectory) && !fileArray[i].StartsWith(commonDirectory, StringComparison.OrdinalIgnoreCase))
                    {
                        commonDirectory = Path.GetDirectoryName(commonDirectory);
                    }
                }
            }

            // Update the text box with the common directory or indicate multiple locations
            var textBox = target switch
            {
                "conversion" => InputFolderTextBox,
                "verification" => VerifyFolderTextBox,
                "extraction" => ExtractInputFolderTextBox,
                _ => null
            };

            textBox?.Text = commonDirectory ?? "(Multiple locations)";

            // Add files to the list (skip duplicates)
            var addedCount = 0;
            var existingPaths = fileList.Select(static f => f.FullPath).ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var file in fileArray)
            {
                if (existingPaths.Contains(file))
                {
                    continue;
                }

                var fileInfo = new FileInfo(file);
                fileList.Add(new Models.FileItem
                {
                    FileName = Path.GetFileName(file),
                    FullPath = file,
                    FileSize = fileInfo.Length,
                    IsSelected = true
                });
                addedCount++;
            }

            // Refresh the DataGrid binding
            switch (target)
            {
                case "conversion":
                    ConversionFilesDataGrid.ItemsSource = _conversionFiles;
                    break;
                case "verification":
                    VerificationFilesDataGrid.ItemsSource = _verificationFiles;
                    break;
                case "extraction":
                    ExtractionFilesDataGrid.ItemsSource = _extractionFiles;
                    break;
            }

            var skippedCount = fileArray.Length - addedCount;
            if (skippedCount > 0)
            {
                LogMessage($"Added {addedCount} file(s) to {target} list. {skippedCount} file(s) were already in the list.");
            }
            else
            {
                LogMessage($"Added {addedCount} file(s) to {target} list.");
            }
        }
        catch (Exception ex)
        {
            LogMessage($"Error adding individual files: {ex.Message}");
            ShowError($"Error adding files: {ex.Message}");
        }
    }

    #endregion
}
