using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using SevenZip;

namespace BatchConvertToRVZ;

public partial class MainWindow : IDisposable
{
    private bool _disposed;
    private Task? _runningTask;
    private bool _dependenciesOk;
    private string? _dolphinToolPath;
    private bool _processSmallerFilesFirst;
    private CancellationTokenSource _cts;
    private readonly UpdateService _updateService;

    private const string GitHubApiUrl = "https://api.github.com/repos/drpetersonfernandes/BatchConvertToRVZ/releases/latest";

    private int _currentDegreeOfParallelismForFiles = 1;

    private const string RvzCompressionMethod = "zstd"; // Default compression method
    private const int RvzCompressionLevel = 5; // Default compression level
    private const int RvzBlockSize = 131072; // Default block size

    // Supported input extensions (Updated to include archives)
    private static readonly string[] AllSupportedInputExtensions = { ".iso", ".zip", ".7z", ".rar" };

    private static readonly string[] ArchiveExtensions = { ".zip", ".7z", ".rar" };

    // Primary target extensions inside archives for RVZ conversion
    private static readonly string[] PrimaryTargetExtensionsInsideArchive = { ".iso", ".gcm", ".wbfs", ".nkit.iso" };

    // Supported extension for verification
    private static readonly string[] RvzExtension = { ".rvz" };

    // Statistics
    private int _totalFilesToProcess;
    private int _successCount;
    private int _failureCount;
    private readonly Stopwatch _operationTimer = new();

    // For Write Speed Calculation
    private const int WriteSpeedUpdateIntervalMs = 1000;

    // Fields for verification move options
    private bool _moveFailedFiles;
    private bool _moveSuccessFiles;


    public MainWindow()
    {
        InitializeComponent();
        _cts = new CancellationTokenSource();

        // The BugReportService is now initialized and managed by the App class.
        _updateService = new UpdateService(GitHubApiUrl);

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
        string? sevenZipDllName;

        try
        {
            dolphinToolExeName = GetDolphinToolExecutableName();
            _dolphinToolPath = Path.Combine(appDirectory, dolphinToolExeName);
            if (!File.Exists(_dolphinToolPath)) missingFiles.Add(dolphinToolExeName);

            sevenZipDllName = GetSevenZipDllName();
            if (!File.Exists(Path.Combine(appDirectory, sevenZipDllName))) missingFiles.Add(sevenZipDllName);
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
            if (dolphinToolExeName != null) LogMessage($"{dolphinToolExeName} found in the application directory.");
            if (sevenZipDllName != null) LogMessage($"{sevenZipDllName} found in the application directory.");
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

    [SuppressMessage("ReSharper", "SwitchExpressionHandlesSomeKnownEnumValuesWithExceptionInDefault")]
    private string GetSevenZipDllName()
    {
        return RuntimeInformation.ProcessArchitecture switch
        {
            Architecture.X64 => "7z_x64.dll",
            Architecture.Arm64 => "7z_arm64.dll",
            _ => throw new PlatformNotSupportedException($"Unsupported architecture: {RuntimeInformation.ProcessArchitecture}")
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

    private async void Window_Closing(object sender, CancelEventArgs e)
    {
        try
        {
            // If a job is in flight, cancel it and keep the window open
            // until the task completes.
            if (_runningTask is { IsCompleted: false })
            {
                e.Cancel = true; // keep window alive
                await _cts.CancelAsync(); // ask workers to stop

                await _runningTask; // wait for graceful stop

                Close(); // now close for real
            }
        }
        catch (Exception ex)
        {
            _ = ReportBugAsync("Error during window closing", ex);
        }
    }

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
            LogMessage($"RVZ Compression: Method={RvzCompressionMethod}, Level={RvzCompressionLevel}, Block Size={RvzBlockSize}");
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
    }

    private void SetControlsState(bool enabled)
    {
        Application.Current.Dispatcher.Invoke(() =>
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
                .Where(file => AllSupportedInputExtensions.Contains(Path.GetExtension(file).ToLowerInvariant()))
                .ToArray();

            // Add sorting by file size if the option is enabled
            if (_processSmallerFilesFirst)
            {
                LogMessage("Sorting files by size (smallest first)...");
                Array.Sort(files, (x, y) =>
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
                    // Do not report this specific error as it's a user issue (archive content), not a bug.
                    if (!extractResult.ErrorMessage.Contains("No supported game image"))
                    {
                        await ReportBugAsync($"Error extracting archive: {Path.GetFileName(inputFile)}", new Exception(extractResult.ErrorMessage));
                    }

                    return false;
                }
            }

            var outputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileToProcess) + ".rvz");

            if (File.Exists(outputFile))
            {
                LogMessage($"Output file already exists, skipping: {Path.GetFileName(outputFile)}");
                return true;
            }

            var success = await ConvertToRvzAsync(dolphinToolPath, fileToProcess, outputFile);

            if (!success || !deleteOriginal) return success;

            if (isArchiveFile)
            {
                // Small delay before deleting original archive
                await Task.Delay(50, _cts.Token);
                await TryDeleteFile(inputFile, $"original archive file: {Path.GetFileName(inputFile)}");
            }
            else
            {
                // Small delay before deleting original ISO
                await Task.Delay(50, _cts.Token);
                await TryDeleteFile(inputFile, $"original ISO file: {Path.GetFileName(inputFile)}");
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
                        await Task.Delay(300 * (i + 1), _cts.Token); // Exponential backoff
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
                await Task.Delay(100, _cts.Token);
                await TryDeleteFile(potentialOutputFile, "partially created RVZ file after error");
            }

            return false;
        }
        finally
        {
            if (isArchiveFile && !string.IsNullOrEmpty(tempDir) && Directory.Exists(tempDir))
            {
                // Small delay before cleaning up temp directory
                await Task.Delay(100, _cts.Token);
                TryDeleteDirectory(tempDir, "temporary extraction directory");
            }
        }
    }

    private async Task<bool> ConvertToRvzAsync(string dolphinToolPath, string inputFile, string outputFile)
    {
        using var process = new Process();
        try
        {
            LogMessage($"Converting '{Path.GetFileName(inputFile)}' to '{Path.GetFileName(outputFile)}'...");

            var arguments = $"convert -i \"{inputFile}\" -o \"{outputFile}\" -f rvz -c {RvzCompressionMethod} -l {RvzCompressionLevel} -b {RvzBlockSize}";

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
                if (_cts.Token.IsCancellationRequested)
                {
                    try
                    {
                        if (!process.HasExited)
                        {
                            // Try graceful termination first
                            process.Kill(true);

                            // Give it a moment to exit gracefully
                            await Task.Delay(150, _cts.Token);

                            // If still running, force kill
                            if (!process.HasExited)
                            {
                                process.Kill();
                                await Task.Delay(100, _cts.Token);
                            }
                        }
                    }
                    catch (Exception killEx)
                    {
                        LogMessage($"Error killing process for {Path.GetFileName(inputFile)}: {killEx.Message}");
                    }

                    _cts.Token.ThrowIfCancellationRequested();
                }

                await Task.Delay(WriteSpeedUpdateIntervalMs, _cts.Token);

                if (process.HasExited || _cts.Token.IsCancellationRequested) break;

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
                await process.WaitForExitAsync(_cts.Token);
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
                    await process.WaitForExitAsync(_cts.Token);
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
                        await Task.Delay(300 * (i + 1), _cts.Token); // Exponential backoff
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
                await Task.Delay(100, _cts.Token);
                await TryDeleteFile(outputFile, "partially created RVZ file after error");
            }

            return false;
        }
        finally
        {
            UpdateWriteSpeedDisplay(0);
            // Process disposal is handled by 'using' statement
        }
    }

    private async Task TryDeleteFile(string filePath, string description)
    {
        try
        {
            if (!File.Exists(filePath)) return;

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
                            await Task.Delay(delay, _cts.Token);
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
                    return;
                }
                catch (IOException ioEx) when (ioEx.Message.Contains("being used by another process") ||
                                               ioEx.Message.Contains("The process cannot access the file"))
                {
                    if (attempt < 9) // Don't log on the last attempt
                    {
                        var delay = 100 * (attempt + 1); // 100ms, 200ms, 300ms, ..., 1000ms
                        LogMessage($"File {Path.GetFileName(filePath)} is still locked. Attempt {attempt + 1}/10. Waiting {delay}ms...");
                        await Task.Delay(delay, _cts.Token);
                    }
                    else
                    {
                        LogMessage($"Failed to delete {description} {Path.GetFileName(filePath)} after 10 attempts: {ioEx.Message}");
                        _ = Task.Run(() => ReportBugAsync($"Failed to delete {description}: {Path.GetFileName(filePath)} after multiple attempts", ioEx));
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
                        await Task.Delay(delay, _cts.Token);
                    }
                    else
                    {
                        LogMessage($"Failed to delete {description} {Path.GetFileName(filePath)} after 10 attempts due to access denied: {authEx.Message}");
                        _ = Task.Run(() => ReportBugAsync($"Access denied when deleting {description}: {Path.GetFileName(filePath)} after multiple attempts", authEx));
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"Failed to delete {description} {Path.GetFileName(filePath)}: {ex.Message}");
                    _ = Task.Run(() => ReportBugAsync($"Failed to delete {description}: {Path.GetFileName(filePath)}", ex));
                    return; // Non-IO exceptions don't benefit from retry
                }
            }
        }
        catch (Exception ex)
        {
            LogMessage($"Error in TryDeleteFile for {description} {filePath}: {ex.Message}");
            _ = Task.Run(() => ReportBugAsync($"Error in TryDeleteFile: {description}", ex));
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

    private async Task<(bool Success, string FilePath, string TempDir, string ErrorMessage)> ExtractArchiveAsync(string archivePath)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        var extension = Path.GetExtension(archivePath).ToLowerInvariant();
        var archiveFileName = Path.GetFileName(archivePath);

        try
        {
            Directory.CreateDirectory(tempDir);

            if (_cts.Token.IsCancellationRequested)
            {
                LogMessage($"Cancellation requested. Skipping extraction for {archiveFileName}.");
                _cts.Token.ThrowIfCancellationRequested();
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
                        // Use 'using' to guarantee disposal of the extractor, even if an exception occurs
                        // (including OperationCanceledException).
                        using var extractor = new SevenZipExtractor(archivePath);

                        // Check for cancellation before starting the potentially long-running operation
                        _cts.Token.ThrowIfCancellationRequested();

                        // Perform the extraction. This is a synchronous operation.
                        // If the token is cancelled during this call, the extractor's
                        // internal state might be inconsistent, but the 'using' block
                        // ensures its finalizer or Dispose method attempts cleanup.
                        extractor.ExtractArchive(tempDir);

                        // Check again after the operation if it was lengthy
                        if (_cts.Token.IsCancellationRequested)
                        {
                            LogMessage($"Extraction of {archiveFileName} completed, but cancellation was requested. Cleaning up.");
                        }

                        _cts.Token.ThrowIfCancellationRequested();
                    }, _cts.Token); // Passing the token here allows Task.Run to respond to cancellation
                    // by throwing an OperationCanceledException, which is caught below.
                    break;

                default: return (false, string.Empty, tempDir, $"Unsupported archive type: {extension}");
            }

            // After successful extraction (or if it wasn't cancelled during),
            // look for the target file.
            var supportedFile = Directory.GetFiles(tempDir, "*.*", SearchOption.AllDirectories)
                .FirstOrDefault(f => PrimaryTargetExtensionsInsideArchive.Contains(Path.GetExtension(f).ToLowerInvariant()));

            return supportedFile != null
                ? (true, supportedFile, tempDir, string.Empty)
                : (false, string.Empty, tempDir, "No supported game image (.iso, .gcm, .wbfs, .nkit.iso) found in archive.");
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
        catch (SevenZipException ex)
        {
            // Handle specific archive errors (corrupt, encrypted, etc.) without sending a bug report.
            var errorMessage = $"Failed to extract archive {archiveFileName}. It may be corrupt, encrypted, or an unsupported format. Error: {ex.Message}";
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
            var match = Regex.Match(progressLine, @"(\d+[\.,]?\d*)%");
            if (!match.Success) return false;

            var percentageStr = match.Groups[1].Value;
            // FIX: Apply replacement before parsing to handle locale-specific decimal separators
            percentageStr = percentageStr.Replace(',', '.');
            if (!double.TryParse(percentageStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var percentage))
                return false;

            LogMessage($"DolphinTool Converting: {percentage:F1}%");
            return true;
        }
        catch (Exception ex)
        {
            LogMessage($"Error parsing DolphinTool progress line '{progressLine}': {ex.Message}");
            return false;
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

            await App.BugReportServiceInstance!.SendBugReportAsync(fullReport.ToString());
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

        _cts?.Cancel();
        _cts?.Dispose();
        _updateService?.Dispose();
        _operationTimer?.Stop();
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
                .Where(file => RvzExtension.Contains(Path.GetExtension(file).ToLowerInvariant()))
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
                    var success = await VerifyRzvFileAsync(dolphinToolPath, inputFile, verifyFolder, moveFailed, moveSuccess);
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

                    var success = await VerifyRzvFileAsync(dolphinToolPath, inputFile, verifyFolder, moveFailed, moveSuccess);
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

    private async Task<bool> VerifyRzvFileAsync(string dolphinToolPath, string inputFile, string baseFolder, bool moveFailed, bool moveSuccess)
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

            await process.WaitForExitAsync(_cts.Token);

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
                if (verificationResult && moveSuccess)
                {
                    await MoveFileToSubfolder(inputFile, baseFolder, "_Success");
                }
                else if (!verificationResult && moveFailed)
                {
                    await MoveFileToSubfolder(inputFile, baseFolder, "_Failed");
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
                await TryDeleteFile(destinationFilePath, $"existing file in {subfolderName} folder");
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
            ProcessingTimeValue.Text = $@"{elapsed:hh\:mm\:ss}";
        });
    }

    private void UpdateWriteSpeedDisplay(double speedInMBps)
    {
        Application.Current.Dispatcher.InvokeAsync(() =>
        {
            WriteSpeedValue.Text = $"{speedInMBps:F1} MB/s";
        });
    }

    private void UpdateProgressDisplay(int current, int total, string currentFileName, string operationVerb)
    {
        var percentage = total == 0 ? 0 : (double)current / total * 100;
        Application.Current.Dispatcher.BeginInvoke(() =>
        {
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
        if (verb.EndsWith('y'))
        {
            return verb[..^1] + "ied";
        }

        return verb.EndsWith('e') ? verb + "d" : verb + "ed";
    }
}