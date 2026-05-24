using BatchConvertToRVZ.services;
using Xunit;

namespace BatchConvertToRVZ.Tests.Services;

public class FileServiceTests
{
    private readonly FileService _fileService = new();

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
}
