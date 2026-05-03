using BatchConvertToRVZ.services;
using Xunit;

namespace BatchConvertToRVZ.Tests.Services;

public class FileServiceTests : IDisposable
{
    private readonly FileService _fileService;
    private readonly List<string> _logs = [];
    private readonly string _testDir;

    public FileServiceTests()
    {
        _fileService = new FileService(msg => _logs.Add(msg));
        _testDir = Path.Combine(Path.GetTempPath(), $"BCRVZ_Test_{Guid.NewGuid()}");
        Directory.CreateDirectory(_testDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testDir))
                Directory.Delete(_testDir, true);
        }
        catch
        {
            // ignore cleanup errors
        }

        GC.SuppressFinalize(this);
    }

    [Theory]
    [InlineData("game.zip", true)]
    [InlineData("game.7z", true)]
    [InlineData("game.rar", true)]
    [InlineData("game.iso", false)]
    [InlineData("game.rvz", false)]
    [InlineData("game", false)]
    public void IsArchiveFileReturnsExpectedResult(string fileName, bool expected)
    {
        var result = _fileService.IsArchiveFile(fileName);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("game.iso", true)]
    [InlineData("game.gcm", true)]
    [InlineData("game.wbfs", true)]
    [InlineData("game.gcz", true)]
    [InlineData("game.wia", true)]
    [InlineData("game.nkit.iso", true)]
    [InlineData("game.zip", true)]
    [InlineData("game.7z", true)]
    [InlineData("game.rar", true)]
    [InlineData("game.rvz", false)]
    [InlineData("game.txt", false)]
    public void IsSupportedInputFileReturnsExpectedResult(string fileName, bool expected)
    {
        var result = _fileService.IsSupportedInputFile(fileName);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("game.rvz", true)]
    [InlineData("game.zip", true)]
    [InlineData("game.7z", true)]
    [InlineData("game.rar", true)]
    [InlineData("game.iso", false)]
    [InlineData("game.txt", false)]
    public void IsSupportedExtractionInputFileReturnsExpectedResult(string fileName, bool expected)
    {
        var result = _fileService.IsSupportedExtractionInputFile(fileName);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("game.rvz", true)]
    [InlineData("game.RVZ", true)]
    [InlineData("game.iso", false)]
    [InlineData("game.zip", false)]
    public void IsRvzFileReturnsExpectedResult(string fileName, bool expected)
    {
        var result = _fileService.IsRvzFile(fileName);
        Assert.Equal(expected, result);
    }

    [Fact]
    public void GetArchiveExtensionsReturnsExpectedValues()
    {
        var extensions = _fileService.GetArchiveExtensions();
        Assert.Equal([".zip", ".7z", ".rar"], extensions);
    }

    [Fact]
    public void GetPrimaryTargetExtensionsInsideArchiveReturnsExpectedValues()
    {
        var extensions = _fileService.GetPrimaryTargetExtensionsInsideArchive();
        Assert.Equal([".iso", ".gcm", ".wbfs", ".rvz", ".gcz", ".wia", ".nkit.iso"], extensions);
    }

    [Fact]
    public void GetRvzExtensionsReturnsExpectedValues()
    {
        var extensions = _fileService.GetRvzExtensions();
        Assert.Equal([".rvz"], extensions);
    }

    [Fact]
    public void GetExtractionInputExtensionsReturnsExpectedValues()
    {
        var extensions = _fileService.GetExtractionInputExtensions();
        Assert.Equal([".rvz", ".zip", ".7z", ".rar"], extensions);
    }

    [Theory]
    [InlineData("game.iso", "game")]
    [InlineData("game.nkit.iso", "game")]
    [InlineData("game.gcm", "game")]
    [InlineData("game.wbfs", "game")]
    [InlineData("game.rvz", "game")]
    [InlineData("game.gcz", "game")]
    [InlineData("game.wia", "game")]
    [InlineData("archive.zip", "archive")]
    public void GetBaseFileNameWithoutGameExtensionReturnsExpectedResult(string fileName, string expected)
    {
        var result = _fileService.GetBaseFileNameWithoutGameExtension(fileName);
        Assert.Equal(expected, result);
    }

    [Fact]
    public void GetFilesFromFolderReturnsMatchingFiles()
    {
        File.WriteAllText(Path.Combine(_testDir, "game1.iso"), "test");
        File.WriteAllText(Path.Combine(_testDir, "game2.rvz"), "test");
        File.WriteAllText(Path.Combine(_testDir, "readme.txt"), "test");
        File.WriteAllText(Path.Combine(_testDir, "game.nkit.iso"), "test");

        var files = _fileService.GetFilesFromFolder(_testDir, [".iso", ".rvz", ".nkit.iso"]);

        Assert.Equal(3, files.Length);
        Assert.Contains(files, static f => f.EndsWith("game1.iso", StringComparison.Ordinal));
        Assert.Contains(files, static f => f.EndsWith("game2.rvz", StringComparison.Ordinal));
        Assert.Contains(files, static f => f.EndsWith("game.nkit.iso", StringComparison.Ordinal));
    }

    [Fact]
    public void GetFilesFromFolderWithRecursiveSearchReturnsNestedFiles()
    {
        var subDir = Path.Combine(_testDir, "sub");
        Directory.CreateDirectory(subDir);
        File.WriteAllText(Path.Combine(_testDir, "game1.iso"), "test");
        File.WriteAllText(Path.Combine(subDir, "game2.iso"), "test");

        var files = _fileService.GetFilesFromFolder(_testDir, [".iso"], true);

        Assert.Equal(2, files.Length);
    }

    [Fact]
    public async Task TryDeleteFileAsyncDeletesExistingFile()
    {
        var filePath = Path.Combine(_testDir, "delete_me.txt");
        File.WriteAllText(filePath, "test");

        var result = await _fileService.TryDeleteFileAsync(filePath, "test file", CancellationToken.None);

        Assert.True(result);
        Assert.False(File.Exists(filePath));
        Assert.Contains(_logs, static l => l.Contains("Deleted test file: delete_me.txt"));
    }

    [Fact]
    public async Task TryDeleteFileAsyncReturnsFalseForMissingFile()
    {
        var filePath = Path.Combine(_testDir, "missing.txt");

        var result = await _fileService.TryDeleteFileAsync(filePath, "test file", CancellationToken.None);

        Assert.False(result);
    }

    [Fact]
    public async Task TryDeleteDirectoryAsyncDeletesExistingDirectory()
    {
        var dirPath = Path.Combine(_testDir, "delete_me");
        Directory.CreateDirectory(dirPath);
        File.WriteAllText(Path.Combine(dirPath, "file.txt"), "test");

        await _fileService.TryDeleteDirectoryAsync(dirPath, "test dir", CancellationToken.None);

        Assert.False(Directory.Exists(dirPath));
        Assert.Contains(_logs, static l => l.Contains("Deleted test dir"));
    }

    [Fact]
    public async Task MoveFileToSubfolderAsyncMovesFileToNewSubfolder()
    {
        var filePath = Path.Combine(_testDir, "game.iso");
        File.WriteAllText(filePath, "test");

        await _fileService.MoveFileToSubfolderAsync(filePath, _testDir, "Converted", CancellationToken.None);

        Assert.False(File.Exists(filePath));
        Assert.True(File.Exists(Path.Combine(_testDir, "Converted", "game.iso")));
    }

    [Fact]
    public async Task MoveFileToSubfolderAsyncRenamesOnCollision()
    {
        var filePath = Path.Combine(_testDir, "game.iso");
        File.WriteAllText(filePath, "test");
        var subDir = Path.Combine(_testDir, "Converted");
        Directory.CreateDirectory(subDir);
        File.WriteAllText(Path.Combine(subDir, "game.iso"), "existing");

        await _fileService.MoveFileToSubfolderAsync(filePath, _testDir, "Converted", CancellationToken.None);

        Assert.False(File.Exists(filePath));
        var movedFiles = Directory.GetFiles(subDir, "game*.iso");
        Assert.Equal(2, movedFiles.Length);
    }

    [Fact]
    public void CreateTempDirectoryCreatesDirectoryWithPrefix()
    {
        var tempDir = _fileService.CreateTempDirectory("TestPrefix_");

        Assert.True(Directory.Exists(tempDir));
        Assert.Contains("TestPrefix_", tempDir);

        // Cleanup
        Directory.Delete(tempDir, true);
    }

    [Theory]
    [InlineData(@"C:\folder\game.iso", "game.iso")]
    [InlineData("game.iso", "game.iso")]
    public void GetFileNameReturnsExpectedResult(string path, string expected)
    {
        Assert.Equal(expected, _fileService.GetFileName(path));
    }

    [Theory]
    [InlineData(@"C:\folder\game.iso", ".iso")]
    public void GetExtensionReturnsExpectedResult(string path, string expected)
    {
        Assert.Equal(expected, _fileService.GetExtension(path));
    }

    [Fact]
    public void CombinePathsReturnsCombinedPath()
    {
        var result = _fileService.CombinePaths("C:\\folder", "sub", "file.iso");
        Assert.Equal(Path.Combine("C:\\folder", "sub", "file.iso"), result);
    }

    [Theory]
    [InlineData("game.iso", ".rvz", "game.rvz")]
    public void ChangeExtensionReturnsExpectedResult(string path, string newExt, string expected)
    {
        Assert.Equal(expected, _fileService.ChangeExtension(path, newExt));
    }
}
