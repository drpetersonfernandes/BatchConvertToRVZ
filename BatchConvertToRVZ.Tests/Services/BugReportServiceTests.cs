using BatchConvertToRVZ.services;
using Xunit;

namespace BatchConvertToRVZ.Tests.Services;

public class BugReportServiceTests
{
    [Fact]
    public void FormatExceptionDetailsFormatsSimpleException()
    {
        var exception = new InvalidOperationException("Something went wrong");

        var result = BugReportService.FormatExceptionDetails(exception);

        Assert.Contains("Type: System.InvalidOperationException", result);
        Assert.Contains("Message: Something went wrong", result);
        Assert.Contains("StackTrace:", result);
    }

    [Fact]
    public void FormatExceptionDetailsFormatsNestedExceptions()
    {
        var inner = new ArgumentException("Inner error");
        var outer = new InvalidOperationException("Outer error", inner);

        var result = BugReportService.FormatExceptionDetails(outer);

        Assert.Contains("Type: System.InvalidOperationException", result);
        Assert.Contains("Message: Outer error", result);
        Assert.Contains("Inner Exception:", result);
        Assert.Contains("Type: System.ArgumentException", result);
        Assert.Contains("Message: Inner error", result);
    }

    [Fact]
    public void FormatExceptionDetailsIncludesSourceInformation()
    {
        var exception = new Exception("Test exception");
        // Source is auto-populated by runtime, but may be empty in tests

        var result = BugReportService.FormatExceptionDetails(exception);

        Assert.Contains("Source:", result);
        Assert.Contains("StackTrace:", result);
    }
}
