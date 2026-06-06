using System.IO.Compression;
using BatchConvertToRVZ.services;
using Xunit;

namespace BatchConvertToRVZ.Tests.Services;

public class ExtractionServiceTests : IDisposable
{
    private readonly FileService _fileService = new();
    private readonly List<string> _logMessages = [];
    private readonly string _tempDir;

    public ExtractionServiceTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "BatchConvertToRVZ_ExtractTests_" + Path.GetRandomFileName());
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, true);
            }
            catch
            {
                /* ignore */
            }
        }

        GC.SuppressFinalize(this);
    }

    private ExtractionService CreateService()
    {
        return new ExtractionService(
            msg => _logMessages.Add(msg), static (_, _) => Task.CompletedTask,
            _fileService);
    }

    private string CreateTestZipArchive(string entryName, byte[]? entryContent = null)
    {
        var archivePath = Path.Combine(_tempDir, $"test_{Path.GetRandomFileName()}.zip");
        using var archive = ZipFile.Open(archivePath, ZipArchiveMode.Create);
        var entry = archive.CreateEntry(entryName);
        using var stream = entry.Open();
        stream.Write(entryContent ?? new byte[100]);
        return archivePath;
    }

    private string CreateCorrupt7ZipArchive()
    {
        var archivePath = Path.Combine(_tempDir, $"test_{Path.GetRandomFileName()}.7z");
        // Write a minimal invalid 7z header
        File.WriteAllBytes(archivePath, [0x37, 0x7A, 0xBC, 0xAF, 0x27, 0x1C, 0x00, 0x00, 0x00, 0x00]);
        return archivePath;
    }

    [Fact]
    public void Get7ZipExecutablePathReturnsPathEndingWith7ZipExe()
    {
        var result = ExtractionService.Get7ZipExecutablePath();

        Assert.EndsWith("7za.exe", result);
    }

    [Fact]
    public void Get7ZipExecutablePathReturnsAbsolutePath()
    {
        var result = ExtractionService.Get7ZipExecutablePath();

        Assert.True(Path.IsPathRooted(result));
    }

    [Fact]
    public void Get7ZipExecutablePathContainsBaseDirectory()
    {
        var result = ExtractionService.Get7ZipExecutablePath();
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;

        Assert.StartsWith(baseDir, result);
    }

    [Fact]
    public void Get7ZipExecutablePathConsistentWithConversionService()
    {
        var conversionResult = ConversionService.Get7ZipExecutablePath();
        var extractionResult = ExtractionService.Get7ZipExecutablePath();

        Assert.Equal(conversionResult, extractionResult);
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncEmptyFilesReturnsImmediately()
    {
        var service = CreateService();

        await service.PerformBatchExtractionAsync(
            "dolphinTool", [], _tempDir, false, "iso", static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

        Assert.Contains("No files selected for extraction.", _logMessages);
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncInvalidFormatReturnsError()
    {
        var service = CreateService();

        await service.PerformBatchExtractionAsync(
            "dolphinTool", ["test.rvz"], _tempDir, false, "invalid", static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

        Assert.Contains(_logMessages, static m => m.Contains("Invalid output format"));
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncCancellationThrowsOperationCanceled()
    {
        var service = CreateService();
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            service.PerformBatchExtractionAsync(
                "dolphinTool", ["test.rvz"], _tempDir, false, "iso", static (_, _, _) => { }, static _ => { }, static _ => { }, cts.Token));
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncUnsupportedFileAttemptsConversion()
    {
        var service = CreateService();
        var filePath = Path.Combine(_tempDir, "test.txt");
        File.WriteAllText(filePath, "not a disc image");

        await service.PerformBatchExtractionAsync(
            @"C:\nonexistent_path\fake_dolphin.exe", [filePath], _tempDir, false, "iso", static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

        // Non-archive files go through DolphinTool conversion path
        Assert.Contains(_logMessages, static m => m.Contains("Converting to ISO"));
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncLogsProcessingMessages()
    {
        var service = CreateService();
        var archivePath = CreateTestZipArchive("game.rvz", new byte[100]);

        await service.PerformBatchExtractionAsync(
            @"C:\nonexistent_path\fake_dolphin.exe", [archivePath], _tempDir, false, "iso", static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

        Assert.Contains(_logMessages, static m => m.Contains("Preparing for batch extraction"));
        Assert.Contains(_logMessages, static m => m.Contains("Processing:"));
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncCorrupt7ZLogsFallbackAttempt()
    {
        var service = CreateService();
        var archivePath = CreateCorrupt7ZipArchive();

        var failureCount = 0;

        await service.PerformBatchExtractionAsync(
            "dolphinTool", [archivePath], _tempDir, false, "iso", static (_, _, _) => { }, static _ => { }, _ => { failureCount++; }, CancellationToken.None);

        Assert.Equal(1, failureCount);
        Assert.Contains(_logMessages, static m => m.Contains("falling back to 7za.exe"));
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncCorrupt7ZReportsWhenBothFail()
    {
        var service = CreateService();
        var archivePath = CreateCorrupt7ZipArchive();

        await service.PerformBatchExtractionAsync(
            "dolphinTool", [archivePath], _tempDir, false, "iso", static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

        Assert.Contains(_logMessages, static m => m.Contains("File may be corrupt") || m.Contains("7za"));
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncZipWithRvzExtractsSuccessfully()
    {
        var service = CreateService();

        var rvzContent = new byte[350_000];
        new Random(42).NextBytes(rvzContent);
        var archivePath = CreateTestZipArchive("game.rvz", rvzContent);

        await service.PerformBatchExtractionAsync(
            @"C:\nonexistent_path\fake_dolphin.exe", [archivePath], _tempDir, false, "iso", static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

        // Extraction should succeed even if DolphinTool conversion fails
        Assert.Contains(_logMessages, static m => m.Contains("Extracted"));
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncZipWithNoRvzReportsError()
    {
        var service = CreateService();
        var archivePath = CreateTestZipArchive("game.iso", new byte[100]);

        var successCount = 0;
        var failureCount = 0;

        await service.PerformBatchExtractionAsync(
            "dolphinTool", [archivePath], _tempDir, false, "iso", static (_, _, _) => { }, _ => { successCount++; }, _ => { failureCount++; }, CancellationToken.None);

        Assert.Equal(0, successCount);
        Assert.Equal(1, failureCount);
        Assert.Contains(_logMessages, static m => m.Contains("No RVZ file found"));
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncMultipleFilesProcessesSequentially()
    {
        var service = CreateService();
        var processedFiles = new List<string>();

        var archive1 = CreateTestZipArchive("game1.rvz", new byte[100]);
        var archive2 = CreateTestZipArchive("game2.rvz", new byte[100]);

        await service.PerformBatchExtractionAsync(
            @"C:\nonexistent_path\fake_dolphin.exe", [archive1, archive2], _tempDir, false, "iso",
            (_, _, name) => processedFiles.Add(name), static _ => { }, static _ => { }, CancellationToken.None);

        Assert.Equal(2, processedFiles.Count);
    }

    [Fact]
    public async Task PerformBatchExtractionAsyncValidFormatsAccepted()
    {
        foreach (var format in new[] { "iso", "wbfs", "gcz", "wia" })
        {
            _logMessages.Clear();
            var service = CreateService();

            var filePath = Path.Combine(_tempDir, $"test_{format}.txt");
            File.WriteAllText(filePath, "not a disc image");

            await service.PerformBatchExtractionAsync(
                @"C:\nonexistent_path\fake_dolphin.exe", [filePath], _tempDir, false, format, static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

            // Should not log "Invalid output format" for valid formats
            Assert.DoesNotContain(_logMessages, static m => m.Contains("Invalid output format"));
        }
    }
}
