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
using SharpCompress.Archives.Zip;
using SharpCompress.Archives.SevenZip;
using SharpCompress.Archives.Rar;

namespace BatchConvertToRVZ;

public partial class MainWindow : IDisposable
{
    private bool _disposed;
    private bool _isClosing;
    private Task? _runningTask;
    private bool _dependenciesOk;
    private string? _dolphinToolPath;
    private bool _processSmallerFilesFirst;
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

    // Statistics
    private int _totalFilesToProcess;
    private int _successCount;
    private int _failureCount;
    private readonly Stopwatch _operationTimer = new();

    // For Write Speed Calculation
    private const int WriteSpeedUpdateIntervalMs = 1000;

    // Thread-safe fields for aggregate write speed tracking during parallel processing
    private long _aggregateTotalBytesWritten;
    private long _aggregateLastTotalBytes;
    private DateTime _aggregateLastCheckTime = DateTime.UtcNow;
    private readonly object _aggregateSpeedLock = new();
    private int _activeConversionCount;

    // Fields for verification move options
    private bool _moveFailedFiles;
    private bool _moveSuccessFiles;


    public MainWindow()
    {
        InitializeComponent();
        _cts = new CancellationTokenSource();

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

    private string GetDolphinToolExecutableName()
    {
        try
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
        catch (Exception ex)
        {
            _ = ReportBugAsync("Error in method GetDolphinToolExecutableName", ex);
        }

        return "DolphinTool.exe";
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

    private readonly SemaphoreSlim _closingSemaphore = new(1, 1);

    private async void Window_Closing(object sender, CancelEventArgs e)
    {
        try
        {
            // Prevent re-entrancy if we're already in the process of closing
            if (_isClosing)
            {
                return;
            }

            // Check if dispatcher is shutting down - if so, just allow the close
            if (Dispatcher.HasShutdownStarted || Dispatcher.HasShutdownFinished)
            {
                return;
            }

            // Use semaphore to prevent concurrent closing attempts
            if (!await _closingSemaphore.WaitAsync(0))
            {
                e.Cancel = true;
                return;
            }

            try
            {
                // If a job is in flight, cancel it and keep the window open
                // until the task completes.
                if (_runningTask is { IsCompleted: false })
                {
                    e.Cancel = true; // keep window alive
                    _isClosing = true; // mark that we're in the closing process

                    try
                    {
                        await _cts.CancelAsync(); // ask workers to stop
                        await _runningTask; // wait for graceful stop
                    }
                    catch (Exception ex)
                    {
                        _isClosing = false; // Reset flag on error so user can try closing again
                        _ = ReportBugAsync("Error during task cancellation while closing", ex);
                        return;
                    }

                    // Close the window on the dispatcher to avoid re-entrancy
                    // Check dispatcher state again before invoking
                    if (!Dispatcher.HasShutdownStarted && !Dispatcher.HasShutdownFinished)
                    {
                        try
                        {
                            await Dispatcher.BeginInvoke(() =>
                            {
                                if (!_disposed && !_isClosing)
                                {
                                    Close();
                                }
                            }, System.Windows.Threading.DispatcherPriority.Normal);
                        }
                        catch (Exception ex)
                        {
                            _isClosing = false;
                            _ = ReportBugAsync("Error during window closing (dispatcher invoke)", ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _isClosing = false;
                _ = ReportBugAsync("Error during window closing", ex);
            }
            finally
            {
                _closingSemaphore.Release();
            }
        }
        catch (Exception ex)
        {
            _ = ReportBugAsync("Error during window closing", ex);
        }
    }

    private const int MaxLogLines = 5000;

    private void LogMessage(string message)
    {
        if (_disposed || Application.Current is null)
        {
            return;
        }

        var timestampedMessage = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";

        Application.Current.Dispatcher.BeginInvoke(() =>
        {
            if (_disposed)
            {
                return;
            }

            LogViewer.AppendText($"{timestampedMessage}{Environment.NewLine}");

            // Trim log to prevent unbounded memory growth
            var lineCount = LogViewer.LineCount;
            if (lineCount > MaxLogLines)
            {
                var linesToRemove = lineCount - MaxLogLines;
                var text = LogViewer.Text;
                var lineIndex = 0;
                for (var i = 0; i < linesToRemove; i++)
                {
                    lineIndex = text.IndexOf('\n', lineIndex);
                    if (lineIndex < 0) break;

                    lineIndex++;
                }

                if (lineIndex > 0)
                {
                    LogViewer.Text = text[lineIndex..];
                }
            }

            LogViewer.ScrollToEnd();
        });
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
            _processSmallerFilesFirst = ProcessSmallerFilesFirstCheckBox.IsChecked ?? true;

            // Update compression settings from UI
            UpdateBlockSizeFromSelection();

            _currentDegreeOfParallelismForFiles = 1;
            if (useParallelFileProcessing && DegreeOfParallelismComboBox.SelectedItem is System.Windows.Controls.ComboBoxItem selectedItem)
            {
                _ = int.TryParse(selectedItem.Content.ToString(), out _currentDegreeOfParallelismForFiles);
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
            LogMessage($"Process smaller files first: {_processSmallerFilesFirst}");

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

        // Show the "Please wait" overlay to inform user that the app is waiting for SharpCompress to finish
        ExtractionOverlayText.Text = "Cancellation requested.\nPlease wait for the current extraction to complete...";
        ExtractionOverlay.Visibility = Visibility.Visible;
    }

    private void SetControlsState(bool enabled)
    {
        Application.Current.Dispatcher.InvokeAsync(() =>
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

            var files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(static file => AllSupportedInputExtensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
                .ToArray();

            // Add sorting by file size if the option is enabled
            if (_processSmallerFilesFirst)
            {
                LogMessage("Sorting files by size (smallest first)...");
                Array.Sort(files, static (x, y) =>
                {
                    try
                    {
                        var xInfo = new FileInfo(x);
                        var yInfo = new FileInfo(y);
                        return xInfo.Length.CompareTo(yInfo.Length);
                    }
                    catch
                    {
                        // If we can't get file info, maintain original order
                        return 0;
                    }
                });
            }

            _totalFilesToProcess = files.Length;
            UpdateStatsDisplay();
            LogMessage($"Found {_totalFilesToProcess} files to process.");
            if (_totalFilesToProcess == 0)
            {
                LogMessage("No supported files (.iso, .zip, .7z, .rar) found in the input folder.");
                return;
            }

            var filesProcessedCount = 0;

            if (useParallelFileProcessing && files.Length > 1)
            {
                // Initialize aggregate speed tracking for parallel processing
                Interlocked.Exchange(ref _aggregateTotalBytesWritten, 0);
                Interlocked.Exchange(ref _aggregateLastTotalBytes, 0);
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

                    var success = await ProcessFileAsync(dolphinToolPath, inputFile, outputFolder, deleteFiles, token);

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

                    var success = await ProcessFileAsync(dolphinToolPath, t, outputFolder, deleteFiles, _cts.Token);

                    if (success)
                    {
                        _successCount++;
                        LogMessage($"Conversion successful: {fileName}");
                    }
                    else
                    {
                        _failureCount++;
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
                    // Do not report this specific error as it's a user issue (archive content), not a bug.
                    if (!extractResult.ErrorMessage.Contains("No supported game image"))
                    {
                        await ReportBugAsync($"Error extracting archive: {Path.GetFileName(inputFile)}", new InvalidOperationException(extractResult.ErrorMessage));
                    }

                    return false;
                }
            }

            cancellationToken.ThrowIfCancellationRequested();

            var outputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileToProcess) + ".rvz");

            if (File.Exists(outputFile))
            {
                LogMessage($"Output file already exists, skipping: {Path.GetFileName(outputFile)}");

                // If user requested to delete original files, delete them even when skipping
                if (deleteOriginal)
                {
                    if (isArchiveFile)
                    {
                        await Task.Delay(50, cancellationToken);
                        await TryDeleteFile(inputFile, $"original archive file: {Path.GetFileName(inputFile)}", cancellationToken);
                    }
                    else
                    {
                        await Task.Delay(50, cancellationToken);
                        await TryDeleteFile(inputFile, $"original ISO file: {Path.GetFileName(inputFile)}", cancellationToken);
                    }
                }

                return true;
            }

            var success = await ConvertToRvzAsync(dolphinToolPath, fileToProcess, outputFile, cancellationToken);

            if (!success || !deleteOriginal) return success;

            if (isArchiveFile)
            {
                // Small delay before deleting original archive
                await Task.Delay(50, cancellationToken);
                await TryDeleteFile(inputFile, $"original archive file: {Path.GetFileName(inputFile)}", cancellationToken);
            }
            else
            {
                // Small delay before deleting original ISO
                await Task.Delay(50, cancellationToken);
                await TryDeleteFile(inputFile, $"original ISO file: {Path.GetFileName(inputFile)}", cancellationToken);
            }

            return success;
        }
        catch (OperationCanceledException)
        {
            LogMessage($"Processing cancelled for {Path.GetFileName(inputFile)}.");

            var potentialOutputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileToProcess) + ".rvz");

            if (File.Exists(potentialOutputFile))
            {
                // Wait a bit longer to ensure the file handle is released
                for (var i = 0; i < 10; i++)
                {
                    try
                    {
                        File.Delete(potentialOutputFile);
                        LogMessage($"Deleted partially created RVZ file after cancellation: {Path.GetFileName(potentialOutputFile)}");
                        break;
                    }
                    catch (IOException)
                    {
                        // File is still locked, wait and retry
                        await Task.Delay(300 * (i + 1), CancellationToken.None); // Exponential backoff - use None as token is already cancelled
                    }
                }
            }

            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error processing file {Path.GetFileName(inputFile)}: {ex.Message}");
            await ReportBugAsync($"Error processing file: {Path.GetFileName(inputFile)}", ex);
            var potentialOutputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileToProcess) + ".rvz");
            if (File.Exists(potentialOutputFile))
            {
                // Small delay before attempting deletion
                await Task.Delay(100, cancellationToken);
                await TryDeleteFile(potentialOutputFile, "partially created RVZ file after error", cancellationToken);
            }

            return false;
        }
        finally
        {
            if (isArchiveFile && !string.IsNullOrEmpty(tempDir) && Directory.Exists(tempDir))
            {
                // Small delay before cleaning up temp directory
                await Task.Delay(100, cancellationToken);
                TryDeleteDirectory(tempDir, "temporary extraction directory");
            }
        }
    }

    private async Task<bool> ConvertToRvzAsync(string dolphinToolPath, string inputFile, string outputFile, CancellationToken cancellationToken)
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
                            await Task.Delay(150, cancellationToken);

                            // If still running, force kill
                            if (!process.HasExited)
                            {
                                process.Kill();
                                await Task.Delay(100, cancellationToken);
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
                            // Only update UI directly if not in parallel mode (single file processing)
                            if (Interlocked.CompareExchange(ref _activeConversionCount, 0, 0) <= 1)
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

            if (File.Exists(outputFile))
            {
                // Wait a bit longer to ensure file handle is released
                for (var i = 0; i < 10; i++)
                {
                    try
                    {
                        File.Delete(outputFile);
                        LogMessage($"Deleted partially created RVZ file after cancellation: {Path.GetFileName(outputFile)}");
                        break;
                    }
                    catch (IOException)
                    {
                        // File is still locked, wait and retry
                        await Task.Delay(300 * (i + 1), CancellationToken.None); // Exponential backoff - use None as token is already cancelled
                    }
                }
            }

            throw;
        }
        catch (Exception ex)
        {
            LogMessage($"Error converting file {Path.GetFileName(inputFile)}: {ex.Message}");
            await ReportBugAsync($"Error converting file: {Path.GetFileName(inputFile)}", ex);
            if (File.Exists(outputFile))
            {
                // Small delay before attempting deletion
                await Task.Delay(100, cancellationToken);
                await TryDeleteFile(outputFile, "partially created RVZ file after error", cancellationToken);
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

            // Try multiple times with increasing delays and different approaches
            for (var attempt = 0; attempt < 10; attempt++) // Increased to 10 attempts
            {
                try
                {
                    // First check if file is locked by trying to open it
                    try
                    {
                        await using (new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                        {
                            // If we can open it exclusively, we can delete it
                        }
                    }
                    catch (IOException)
                    {
                        // File is locked, wait and retry
                        if (attempt < 9) // Don't log on the last attempt
                        {
                            var delay = 50 * (attempt + 1); // 50ms, 100ms, 150ms, ..., 500ms
                            LogMessage($"File {Path.GetFileName(filePath)} is still locked. Attempt {attempt + 1}/10. Waiting {delay}ms...");
                            await Task.Delay(delay, CancellationToken.None);
                            continue;
                        }
                    }

                    // Try to remove read-only attribute if it exists
                    var attributes = File.GetAttributes(filePath);
                    if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        File.SetAttributes(filePath, attributes & ~FileAttributes.ReadOnly);
                    }

                    // Finally, try to delete the file
                    File.Delete(filePath);
                    LogMessage($"Deleted {description}: {Path.GetFileName(filePath)}");
                    return true;
                }
                catch (IOException ioEx) when ((ioEx.HResult & 0xFFFF) == 32 || // ERROR_SHARING_VIOLATION
                                               (ioEx.HResult & 0xFFFF) == 33) // ERROR_LOCK_VIOLATION
                {
                    if (attempt < 9) // Don't log on the last attempt
                    {
                        var delay = 100 * (attempt + 1); // 100ms, 200ms, 300ms, ..., 1000ms
                        LogMessage($"File {Path.GetFileName(filePath)} is still locked. Attempt {attempt + 1}/10. Waiting {delay}ms...");
                        await Task.Delay(delay, CancellationToken.None);
                    }
                    else
                    {
                        LogMessage($"Failed to delete {description} {Path.GetFileName(filePath)} after 10 attempts: {ioEx.Message}");
                        _ = Task.Run(() => ReportBugAsync($"Failed to delete {description}: {Path.GetFileName(filePath)} after multiple attempts", ioEx), cancellationToken);
                        return false;
                    }
                }
                catch (UnauthorizedAccessException authEx)
                {
                    // Try to reset file attributes
                    try
                    {
                        File.SetAttributes(filePath, FileAttributes.Normal);
                    }
                    catch
                    {
                        // Ignore errors when trying to reset attributes
                    }

                    if (attempt < 9)
                    {
                        var delay = 150 * (attempt + 1);
                        LogMessage($"Access denied for {Path.GetFileName(filePath)}. Attempt {attempt + 1}/10. Waiting {delay}ms...");
                        await Task.Delay(delay, CancellationToken.None);
                    }
                    else
                    {
                        LogMessage($"Failed to delete {description} {Path.GetFileName(filePath)} after 10 attempts due to access denied: {authEx.Message}");
                        _ = Task.Run(() => ReportBugAsync($"Access denied when deleting {description}: {Path.GetFileName(filePath)} after multiple attempts", authEx), cancellationToken);
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"Failed to delete {description} {Path.GetFileName(filePath)}: {ex.Message}");
                    _ = Task.Run(() => ReportBugAsync($"Failed to delete {description}: {Path.GetFileName(filePath)}", ex), cancellationToken);
                    return false; // Non-IO exceptions don't benefit from retry
                }
            }

            // If we exhausted all attempts without returning, deletion failed
            return false;
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

    private async Task<(bool Success, string FilePath, string TempDir, string ErrorMessage)> ExtractArchiveAsync(string archivePath, CancellationToken cancellationToken)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        var extension = Path.GetExtension(archivePath);
        var archiveFileName = Path.GetFileName(archivePath);

        try
        {
            Directory.CreateDirectory(tempDir);

            if (cancellationToken.IsCancellationRequested)
            {
                LogMessage($"Cancellation requested. Skipping extraction for {archiveFileName}.");
                cancellationToken.ThrowIfCancellationRequested();
            }

            LogMessage($"Extracting {archiveFileName} to temporary directory: {tempDir}");

            switch (extension)
            {
                case ".zip":
                case ".7z":
                case ".rar":
                    // Wrap the synchronous extraction in a Task to avoid blocking the UI thread,
                    // but ensure cancellation is handled correctly and resources are disposed.
                    await Task.Run(() =>
                    {
                        // Check for cancellation before starting the potentially long-running operation
                        cancellationToken.ThrowIfCancellationRequested();

                        // Open the archive using SharpCompress based on extension
                        IArchive? archive = null;
                        try
                        {
                            archive = extension switch
                            {
                                ".zip" => ZipArchive.OpenArchive(archivePath),
                                ".7z" => SevenZipArchive.OpenArchive(archivePath),
                                ".rar" => RarArchive.OpenArchive(archivePath),
                                _ => throw new InvalidOperationException($"Unsupported archive type: {extension}")
                            };

                            // Extract all entries to the temp directory
                            // Get full path of tempDir for ZipSlip vulnerability protection
                            var tempDirFullPath = Path.GetFullPath(tempDir);
                            if (!tempDirFullPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                            {
                                tempDirFullPath += Path.DirectorySeparatorChar;
                            }

                            foreach (var entry in archive.Entries.Where(static e => !e.IsDirectory && !string.IsNullOrEmpty(e.Key)))
                            {
                                cancellationToken.ThrowIfCancellationRequested();

                                if (entry.Key != null)
                                {
                                    // Filter by extension before extracting to avoid wasting disk I/O on unwanted files
                                    var entryExtension = Path.GetExtension(entry.Key);
                                    if (!PrimaryTargetExtensionsInsideArchive.Contains(entryExtension, StringComparer.OrdinalIgnoreCase))
                                    {
                                        continue;
                                    }

                                    var entryPath = Path.Combine(tempDir, entry.Key);
                                    var entryFullPath = Path.GetFullPath(entryPath);

                                    // ZipSlip protection: ensure the entry path is within the temp directory
                                    if (!entryFullPath.StartsWith(tempDirFullPath, StringComparison.OrdinalIgnoreCase))
                                    {
                                        LogMessage($"Skipping potentially malicious archive entry with directory traversal: {entry.Key}");
                                        continue;
                                    }

                                    var entryDir = Path.GetDirectoryName(entryPath);

                                    if (!string.IsNullOrEmpty(entryDir) && !Directory.Exists(entryDir))
                                    {
                                        Directory.CreateDirectory(entryDir);
                                    }

                                    entry.WriteToFile(entryPath);
                                }
                            }
                        }
                        finally
                        {
                            (archive as IDisposable)?.Dispose();
                        }

                        // Check again after the operation if it was lengthy
                        if (cancellationToken.IsCancellationRequested)
                        {
                            LogMessage($"Extraction of {archiveFileName} completed, but cancellation was requested. Cleaning up.");
                        }

                        cancellationToken.ThrowIfCancellationRequested();
                    }, cancellationToken); // Passing the token here allows Task.Run to respond to cancellation
                    // by throwing an OperationCanceledException, which is caught below.
                    break;

                default: return (false, string.Empty, tempDir, $"Unsupported archive type: {extension}");
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
            return (false, string.Empty, tempDir, "No supported game image (.iso, .gcm, .wbfs, .nkit.iso) found in archive.");
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
        catch (InvalidOperationException ex) when (ex.Message.Contains("archive") || ex.Message.Contains("password") || ex.Message.Contains("encrypted"))
        {
            // Handle specific archive errors (corrupt, encrypted, etc.) without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. It may be corrupt, encrypted, or an unsupported format. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, tempDir, errorMessage);
        }
        catch (IOException ex) when (ex.Message.Contains("corrupt") || ex.Message.Contains("invalid"))
        {
            // Handle specific archive errors (corrupt, etc.) without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. The archive may be corrupt. Error: {ex.Message}";
            LogMessage(errorMessage);

            // Clean up the temporary directory on failure
            TryDeleteDirectory(tempDir, $"failed extraction directory for {archiveFileName}");

            // Return a failure result with the detailed error message
            return (false, string.Empty, tempDir, errorMessage);
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
            return (false, string.Empty, tempDir, errorMessage);
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
            return (false, string.Empty, tempDir, $"Exception during extraction: {ex.Message}");
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
                        return numberStr.Replace(',', '.');
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
            fullReport.AppendLine("Application: BatchConvertToRVZ");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $"Version: {GetType().Assembly.GetName().Version}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $"OS: {Environment.OSVersion}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $".NET Version: {Environment.Version}");
            fullReport.AppendLine(CultureInfo.InvariantCulture, $"Architecture: {RuntimeInformation.ProcessArchitecture}");
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

            if (LogViewer != null)
            {
                var logContent = string.Empty;
                await Application.Current.Dispatcher.InvokeAsync(() => logContent = LogViewer.Text);
                if (!string.IsNullOrEmpty(logContent))
                {
                    fullReport.AppendLine().AppendLine("=== Application Log ===").Append(logContent);
                }
            }

            if (App.BugReportServiceInstance != null) await App.BugReportServiceInstance.SendBugReportAsync(fullReport.ToString());
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
        Application.Current.Dispatcher.InvokeAsync(() =>
        {
            ProgressBar.Value = 0; // Reset progress
            ProgressBar.Maximum = 1; // Ensure the maximum is not zero when idle
            ProgressText.Text = "Ready."; // Set a default idle message
        });
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
        _closingSemaphore.Dispose();
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

            if (string.IsNullOrEmpty(verifyFolder))
            {
                LogMessage("Error: No verification folder selected.");
                ShowError("Please select the folder containing RVZ files to verify.");
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
                LogMessage("No RVZ files (.rvz) found in the selected folder.");
                return;
            }

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
                        _successCount++;
                    }
                    else
                    {
                        _failureCount++;
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

            await process.WaitForExitAsync(token);

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
        Interlocked.Exchange(ref _aggregateLastTotalBytes, 0);
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
        Application.Current.Dispatcher.InvokeAsync(() =>
        {
            TotalFilesValue.Text = _totalFilesToProcess.ToString(CultureInfo.InvariantCulture);
            SuccessValue.Text = _successCount.ToString(CultureInfo.InvariantCulture);
            FailedValue.Text = _failureCount.ToString(CultureInfo.InvariantCulture);
        });
    }

    private void UpdateProcessingTimeDisplay()
    {
        var elapsed = _operationTimer.Elapsed;
        Application.Current.Dispatcher.InvokeAsync(() =>
        {
            ProcessingTimeValue.Text = $"{(int)elapsed.TotalHours:D2}:{elapsed:mm\\:ss}";
        });
    }

    private void UpdateWriteSpeedDisplay(double speedInMBps)
    {
        Application.Current.Dispatcher.InvokeAsync(() =>
        {
            WriteSpeedValue.Text = $"{speedInMBps:F1} MB/s";
        });
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

                // Only update if there are active conversions
                if (Interlocked.CompareExchange(ref _activeConversionCount, 0, 0) == 0)
                    continue;

                lock (_aggregateSpeedLock)
                {
                    var currentTime = DateTime.UtcNow;
                    var timeDelta = currentTime - _aggregateLastCheckTime;

                    if (timeDelta.TotalSeconds > 0)
                    {
                        var currentTotalBytes = Interlocked.Read(ref _aggregateTotalBytesWritten);
                        var bytesDelta = currentTotalBytes - _aggregateLastTotalBytes;
                        var speedInMBps = (bytesDelta / timeDelta.TotalSeconds) / (1024.0 * 1024.0);

                        UpdateWriteSpeedDisplay(speedInMBps);

                        _aggregateLastTotalBytes = currentTotalBytes;
                        _aggregateLastCheckTime = currentTime;
                    }
                }
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
        Application.Current.Dispatcher.Invoke(() =>
        {
            var percentage = total == 0 ? 0 : (double)current / total * 100;
            ProgressText.Text = $"{operationVerb} file {current} of {total}: {currentFileName} ({percentage:F1}%)";
            ProgressBar.Value = current;
            ProgressBar.Maximum = Math.Max(total, 1);
        });
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
}