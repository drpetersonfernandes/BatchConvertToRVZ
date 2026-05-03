using BatchConvertToRVZ.services;
using Xunit;

namespace BatchConvertToRVZ.Tests.Services;

public class UpdateServiceTests
{
    [Theory]
    [InlineData("1.0", "1.0.0.0")]
    [InlineData("1.2.3", "1.2.3.0")]
    [InlineData("1.2.3.4", "1.2.3.4")]
    [InlineData("0.0", "0.0.0.0")]
    public void NormalizeVersionReturnsFourComponentVersion(string input, string expected)
    {
        var inputVersion = Version.Parse(input);
        var result = UpdateService.NormalizeVersion(inputVersion);

        Assert.NotNull(result);
        Assert.Equal(expected, result.ToString());
    }

    [Fact]
    public void NormalizeVersionNullReturnsNull()
    {
        var result = UpdateService.NormalizeVersion(null);
        Assert.Null(result);
    }

    [Theory]
    [InlineData("v1.0.0", "1.0.0")]
    [InlineData("V1.2.3", "1.2.3")]
    [InlineData("release_1.8.0", "1.8.0")]
    [InlineData("1.7", "1.7")]
    [InlineData("v1.2.3-beta.1", "1.2.3")]
    [InlineData("v1.2.3.4", "1.2.3.4")]
    [InlineData("v2.1.0", "2.1.0")]
    public void ParseVersionFromTagReturnsExpectedVersion(string tag, string expected)
    {
        var result = UpdateService.ParseVersionFromTag(tag);

        Assert.NotNull(result);
        Assert.Equal(expected, result.ToString());
    }

    [Theory]
    [InlineData("")]
    [InlineData("v")]
    [InlineData("invalid")]
    [InlineData("version-one")]
    [InlineData(null)]
    public void ParseVersionFromTagInvalidTagsReturnNull(string? tag)
    {
        if (tag != null)
        {
            var result = UpdateService.ParseVersionFromTag(tag);
            Assert.Null(result);
        }
    }

    [Theory]
    [InlineData("1.0.0.0", "1.0.0.0", false)]
    [InlineData("1.0.0.0", "1.0.0.1", true)]
    [InlineData("1.0.0.0", "1.0.1.0", true)]
    [InlineData("1.0.0.0", "1.1.0.0", true)]
    [InlineData("1.0.0.0", "2.0.0.0", true)]
    [InlineData("1.8.1.0", "1.8.0.0", false)]
    public void VersionComparisonWorksAsExpected(string current, string latest, bool expectedUpdate)
    {
        var currentVersion = UpdateService.NormalizeVersion(Version.Parse(current));
        var latestVersion = UpdateService.NormalizeVersion(Version.Parse(latest));

        var isUpdateAvailable = latestVersion > currentVersion;
        Assert.Equal(expectedUpdate, isUpdateAvailable);
    }
}
