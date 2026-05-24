using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using BatchConvertToRVZ.Models;
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
    [InlineData("release-1.8.0", "1.8.0")]
    [InlineData("release-2.0", "2.0")]
    [InlineData("v1.2.3-rc.2", "1.2.3")]
    [InlineData("V2.1.0-ALPHA", "2.1.0")]
    public void ParseVersionFromTagHyphenAndSuffixFormats(string tag, string expected)
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

    [Fact]
    public async Task CheckForUpdatesAsyncUpdateAvailableReturnsTrueAndRelease()
    {
        var release = new GitHubRelease
        {
            TagName = "v99.0.0",
            Name = "v99.0.0",
            HtmlUrl = "https://github.com/test/releases/tag/v99.0.0",
            Body = "Test release",
            Prerelease = false,
            Draft = false
        };

        using var handler = new FakeHttpMessageHandler(CreateJsonResponse(release));
        using var service = new UpdateService("https://api.github.com/test", handler);

        var (isUpdateAvailable, latestRelease) = await service.CheckForUpdatesAsync();

        Assert.True(isUpdateAvailable);
        Assert.NotNull(latestRelease);
        Assert.Equal("v99.0.0", latestRelease.TagName);
        Assert.Equal("https://github.com/test/releases/tag/v99.0.0", latestRelease.HtmlUrl);
    }

    [Fact]
    public async Task CheckForUpdatesAsyncNoUpdateReturnsFalseAndNull()
    {
        var release = new GitHubRelease
        {
            TagName = "v0.1.0",
            Name = "v0.1.0",
            HtmlUrl = "https://github.com/test/releases/tag/v0.1.0",
            Body = "Old release",
            Prerelease = false,
            Draft = false
        };

        using var handler = new FakeHttpMessageHandler(CreateJsonResponse(release));
        using var service = new UpdateService("https://api.github.com/test", handler);

        var (isUpdateAvailable, latestRelease) = await service.CheckForUpdatesAsync();

        Assert.False(isUpdateAvailable);
        Assert.Null(latestRelease);
    }

    [Fact]
    public async Task CheckForUpdatesAsyncPrereleaseReturnsFalseAndNull()
    {
        var release = new GitHubRelease
        {
            TagName = "v99.0.0-beta.1",
            Name = "v99.0.0-beta.1",
            HtmlUrl = "https://github.com/test/releases/tag/v99.0.0-beta.1",
            Body = "Prerelease",
            Prerelease = true,
            Draft = false
        };

        using var handler = new FakeHttpMessageHandler(CreateJsonResponse(release));
        using var service = new UpdateService("https://api.github.com/test", handler);

        var (isUpdateAvailable, latestRelease) = await service.CheckForUpdatesAsync();

        Assert.False(isUpdateAvailable);
        Assert.Null(latestRelease);
    }

    [Fact]
    public async Task CheckForUpdatesAsyncDraftReturnsFalseAndNull()
    {
        var release = new GitHubRelease
        {
            TagName = "v99.0.0",
            Name = "v99.0.0",
            HtmlUrl = "https://github.com/test/releases/tag/v99.0.0",
            Body = "Draft",
            Prerelease = false,
            Draft = true
        };

        using var handler = new FakeHttpMessageHandler(CreateJsonResponse(release));
        using var service = new UpdateService("https://api.github.com/test", handler);

        var (isUpdateAvailable, latestRelease) = await service.CheckForUpdatesAsync();

        Assert.False(isUpdateAvailable);
        Assert.Null(latestRelease);
    }

    [Fact]
    public async Task CheckForUpdatesAsyncNullResponseReturnsFalseAndNull()
    {
        using var handler = new FakeHttpMessageHandler(new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent("null", MediaTypeHeaderValue.Parse("application/json"))
        });
        using var service = new UpdateService("https://api.github.com/test", handler);

        var (isUpdateAvailable, latestRelease) = await service.CheckForUpdatesAsync();

        Assert.False(isUpdateAvailable);
        Assert.Null(latestRelease);
    }

    [Fact]
    public async Task CheckForUpdatesAsyncNetworkErrorRethrows()
    {
        using var handler = new FakeHttpMessageHandler(new HttpRequestException("Network error"));
        using var service = new UpdateService("https://api.github.com/test", handler);

        await Assert.ThrowsAsync<HttpRequestException>(service.CheckForUpdatesAsync);
    }

    private static HttpResponseMessage CreateJsonResponse<T>(T obj)
    {
        var json = JsonSerializer.Serialize(obj);
        return new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(json, MediaTypeHeaderValue.Parse("application/json"))
        };
    }

    private sealed class FakeHttpMessageHandler : HttpMessageHandler
    {
        private readonly Func<HttpRequestMessage, HttpResponseMessage> _handler;

        public FakeHttpMessageHandler(HttpResponseMessage response)
            : this(_ => response)
        {
        }

        public FakeHttpMessageHandler(Exception exception)
            : this(_ => throw exception)
        {
        }

        private FakeHttpMessageHandler(Func<HttpRequestMessage, HttpResponseMessage> handler)
        {
            _handler = handler;
        }

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            return Task.FromResult(_handler(request));
        }
    }
}
