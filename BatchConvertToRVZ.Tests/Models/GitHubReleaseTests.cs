using System.Text.Json;
using BatchConvertToRVZ.Models;
using Xunit;

namespace BatchConvertToRVZ.Tests.Models;

public class GitHubReleaseTests
{
    [Fact]
    public void DeserializeFullReleaseJsonReturnsExpectedValues()
    {
        const string json = """
                            {
                                "tag_name": "v2.1.0",
                                "html_url": "https://github.com/user/repo/releases/tag/v2.1.0",
                                "name": "Release 2.1.0",
                                "body": "Bug fixes and improvements",
                                "prerelease": false,
                                "draft": false
                            }
                            """;

        var release = JsonSerializer.Deserialize<GitHubRelease>(json);

        Assert.NotNull(release);
        Assert.Equal("v2.1.0", release.TagName);
        Assert.Equal("https://github.com/user/repo/releases/tag/v2.1.0", release.HtmlUrl);
        Assert.Equal("Release 2.1.0", release.Name);
        Assert.Equal("Bug fixes and improvements", release.Body);
        Assert.False(release.Prerelease);
        Assert.False(release.Draft);
    }

    [Fact]
    public void DeserializePrereleaseJsonReturnsExpectedValues()
    {
        const string json = """
                            {
                                "tag_name": "v2.2.0-beta.1",
                                "html_url": "https://github.com/user/repo/releases/tag/v2.2.0-beta.1",
                                "name": "Beta Release",
                                "body": "Beta features",
                                "prerelease": true,
                                "draft": false
                            }
                            """;

        var release = JsonSerializer.Deserialize<GitHubRelease>(json);

        Assert.NotNull(release);
        Assert.True(release.Prerelease);
        Assert.False(release.Draft);
    }

    [Fact]
    public void DeserializeDraftReleaseJsonReturnsExpectedValues()
    {
        const string json = """
                            {
                                "tag_name": "v3.0.0",
                                "html_url": "https://github.com/user/repo/releases/tag/v3.0.0",
                                "name": "Draft Release",
                                "body": "Work in progress",
                                "prerelease": false,
                                "draft": true
                            }
                            """;

        var release = JsonSerializer.Deserialize<GitHubRelease>(json);

        Assert.NotNull(release);
        Assert.False(release.Prerelease);
        Assert.True(release.Draft);
    }

    [Fact]
    public void DeserializeMinimalJsonReturnsDefaultValues()
    {
        const string json = "{}";

        var release = JsonSerializer.Deserialize<GitHubRelease>(json);

        Assert.NotNull(release);
        Assert.Equal(string.Empty, release.TagName);
        Assert.Equal(string.Empty, release.HtmlUrl);
        Assert.Equal(string.Empty, release.Name);
        Assert.Equal(string.Empty, release.Body);
        Assert.False(release.Prerelease);
        Assert.False(release.Draft);
    }

    [Fact]
    public void DeserializeExtraFieldIsGraceful()
    {
        const string json = """
                            {
                                "tag_name": "v1.0.0",
                                "extra_field": "should be ignored"
                            }
                            """;

        var release = JsonSerializer.Deserialize<GitHubRelease>(json);

        Assert.NotNull(release);
        Assert.Equal("v1.0.0", release.TagName);
    }
}
