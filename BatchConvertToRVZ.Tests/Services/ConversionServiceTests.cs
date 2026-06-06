using System.IO.Compression;
using BatchConvertToRVZ.services;
using Xunit;

namespace BatchConvertToRVZ.Tests.Services;

public class ConversionServiceTests : IDisposable
{
    private readonly FileService _fileService = new();
    private readonly List<string> _logMessages = [];
    private readonly string _tempDir;

    public ConversionServiceTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "BatchConvertToRVZ_Tests_" + Path.GetRandomFileName());
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

    private ConversionService CreateService()
    {
        return new ConversionService(
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
        File.WriteAllBytes(archivePath, [0x37, 0x7A, 0xBC, 0xAF, 0x27, 0x1C, 0x00, 0x00, 0x00, 0x00]);
        return archivePath;
    }

    [Fact]
    public void Get7ZipExecutablePathReturnsPathEndingWith7ZipExe()
    {
        var result = ConversionService.Get7ZipExecutablePath();

        Assert.EndsWith("7za.exe", result);
    }

    [Fact]
    public void Get7ZipExecutablePathReturnsAbsolutePath()
    {
        var result = ConversionService.Get7ZipExecutablePath();

        Assert.True(Path.IsPathRooted(result));
    }

    [Fact]
    public void Get7ZipExecutablePathContainsBaseDirectory()
    {
        var result = ConversionService.Get7ZipExecutablePath();
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;

        Assert.StartsWith(baseDir, result);
    }

    [Fact]
    public async Task PerformBatchConversionAsyncEmptyFilesReturnsImmediately()
    {
        var service = CreateService();

        await service.PerformBatchConversionAsync(
            "dolphinTool", [], _tempDir, false, "zstd", 5, 131072, static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

        Assert.Contains("No files selected for conversion.", _logMessages);
    }

    [Fact]
    public async Task PerformBatchConversionAsyncCancellationThrowsOperationCanceled()
    {
        var service = CreateService();
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            service.PerformBatchConversionAsync(
                "dolphinTool", ["test.7z"], _tempDir, false, "zstd", 5, 131072, static (_, _, _) => { }, static _ => { }, static _ => { }, cts.Token));
    }

    [Fact]
    public async Task PerformBatchConversionAsyncUnsupportedFileReturnsFailure()
    {
        var service = CreateService();
        var filePath = Path.Combine(_tempDir, "test.txt");
        File.WriteAllText(filePath, "not a disc image");

        var successCount = 0;
        var failureCount = 0;

        await service.PerformBatchConversionAsync(
            @"C:\nonexistent_path\fake_dolphin.exe", [filePath], _tempDir, false, "zstd", 5, 131072, static (_, _, _) => { }, _ => { successCount++; }, _ => { failureCount++; }, CancellationToken.None);

        Assert.Equal(0, successCount);
        Assert.Equal(1, failureCount);
    }

    [Fact]
    public async Task PerformBatchConversionAsyncCorrupt7ZLogsFallbackAttempt()
    {
        var service = CreateService();
        var archivePath = CreateCorrupt7ZipArchive();

        var failureCount = 0;

        await service.PerformBatchConversionAsync(
            "dolphinTool", [archivePath], _tempDir, false, "zstd", 5, 131072, static (_, _, _) => { }, static _ => { }, _ => { failureCount++; }, CancellationToken.None);

        Assert.Equal(1, failureCount);
        Assert.Contains(_logMessages, static m => m.Contains("falling back to 7za.exe"));
    }

    [Fact]
    public async Task PerformBatchConversionAsyncCorrupt7ZReportsBugWhenBothFail()
    {
        var service = CreateService();
        var archivePath = CreateCorrupt7ZipArchive();

        await service.PerformBatchConversionAsync(
            "dolphinTool", [archivePath], _tempDir, false, "zstd", 5, 131072, static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

        Assert.Contains(_logMessages, static m => m.Contains("File may be corrupt") || m.Contains("7za"));
    }

    [Fact]
    public async Task PerformBatchConversionAsyncZipWithIsoExtractsSuccessfully()
    {
        var service = CreateService();

        // Create a minimal ISO-like file (just enough for SharpCompress to extract)
        var isoContent = new byte[350_000]; // Small ISO-like file
        new Random(42).NextBytes(isoContent);
        var archivePath = CreateTestZipArchive("game.iso", isoContent);

        var failureCount = 0;

        await service.PerformBatchConversionAsync(
            @"C:\nonexistent_path\fake_dolphin.exe", [archivePath], _tempDir, false, "zstd", 5, 131072, static (_, _, _) => { }, static _ => { }, _ => { failureCount++; }, CancellationToken.None);

        // Extraction should succeed, conversion fails due to missing tool
        Assert.Contains(_logMessages, static m => m.Contains("Extracted"));
        Assert.Equal(1, failureCount);
    }

    [Fact]
    public async Task PerformBatchConversionAsyncZipWithRvzCopiesDirectly()
    {
        var service = CreateService();

        var rvzContent = new byte[100];
        new Random(42).NextBytes(rvzContent);
        var archivePath = CreateTestZipArchive("game.rvz", rvzContent);

        var successCount = 0;

        await service.PerformBatchConversionAsync(
            @"C:\nonexistent_path\fake_dolphin.exe", [archivePath], _tempDir, false, "zstd", 5, 131072, static (_, _, _) => { }, _ => { successCount++; }, static _ => { }, CancellationToken.None);

        Assert.Equal(1, successCount);
        Assert.Contains(_logMessages, static m => m.Contains("Found RVZ file inside archive"));
    }

    [Fact]
    public async Task PerformBatchConversionAsyncZipWithNoSupportedFileReportsError()
    {
        var service = CreateService();
        var archivePath = CreateTestZipArchive("readme.txt", "hello world"u8.ToArray());

        var successCount = 0;
        var failureCount = 0;

        await service.PerformBatchConversionAsync(
            "dolphinTool", [archivePath], _tempDir, false, "zstd", 5, 131072, static (_, _, _) => { }, _ => { successCount++; }, _ => { failureCount++; }, CancellationToken.None);

        Assert.Equal(0, successCount);
        Assert.Equal(1, failureCount);
        Assert.Contains(_logMessages, static m => m.Contains("No supported disc image found"));
    }

    [Fact]
    public async Task PerformBatchConversionAsyncMultipleFilesProcessesSequentially()
    {
        var service = CreateService();
        var processedFiles = new List<string>();

        var archive1 = CreateTestZipArchive("game1.iso", new byte[100]);
        var archive2 = CreateTestZipArchive("game2.iso", new byte[100]);

        await service.PerformBatchConversionAsync(
            @"C:\nonexistent_path\fake_dolphin.exe", [archive1, archive2], _tempDir, false, "zstd", 5, 131072,
            (_, _, name) => processedFiles.Add(name), static _ => { }, static _ => { }, CancellationToken.None);

        Assert.Equal(2, processedFiles.Count);
    }

    [Fact]
    public async Task PerformBatchConversionAsyncLogsProcessingMessages()
    {
        var service = CreateService();
        var archivePath = CreateTestZipArchive("game.iso", new byte[100]);

        await service.PerformBatchConversionAsync(
            @"C:\nonexistent_path\fake_dolphin.exe", [archivePath], _tempDir, false, "zstd", 5, 131072, static (_, _, _) => { }, static _ => { }, static _ => { }, CancellationToken.None);

        Assert.Contains(_logMessages, static m => m.Contains("Preparing for batch conversion"));
        Assert.Contains(_logMessages, static m => m.Contains("Processing:"));
    }
}
