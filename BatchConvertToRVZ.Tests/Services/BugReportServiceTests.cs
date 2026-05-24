using BatchConvertToRVZ.services;
using Xunit;

namespace BatchConvertToRVZ.Tests.Services;

public class BugReportServiceTests
{
    [Fact]
    public void GetWindowsVersionReturnsNonEmptyString()
    {
        var result = BugReportService.GetWindowsVersion();

        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void GetWindowsVersionReturnsStringContainingWindows()
    {
        var result = BugReportService.GetWindowsVersion();

        Assert.Contains("Windows", result);
    }
}
