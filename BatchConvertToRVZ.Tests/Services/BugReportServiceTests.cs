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
        var exception = new InvalidOperationException("Test exception");
        // Source is auto-populated by runtime, but may be empty in tests

        var result = BugReportService.FormatExceptionDetails(exception);

        Assert.Contains("Source:", result);
        Assert.Contains("StackTrace:", result);
    }

    [Fact]
    public void FormatExceptionDetailsHandlesDeeplyNestedExceptions()
    {
        var inner3 = new NotSupportedException("Param was null");
        var inner2 = new InvalidOperationException("Mid error", inner3);
        var inner1 = new IOException("IO error", inner2);
        var outer = new AggregateException("Agg error", inner1);

        var result = BugReportService.FormatExceptionDetails(outer);

        Assert.Contains("Type: System.AggregateException", result);
        Assert.Contains("Type: System.IO.IOException", result);
        Assert.Contains("Type: System.InvalidOperationException", result);
        Assert.Contains("Type: System.NotSupportedException", result);
        Assert.Contains("Message: Param was null", result);
        Assert.Contains("Message: Agg error", result);
        Assert.Contains("Message: IO error", result);
        Assert.Contains("Message: Mid error", result);
    }

    [Fact]
    public void FormatExceptionDetailsWithNullPropertiesDoesNotThrow()
    {
        var ex = CreateMinimalException();
        var result = BugReportService.FormatExceptionDetails(ex);

        Assert.NotNull(result);
        Assert.NotEmpty(result);
        return;

        // Create an exception with minimal info by throwing and catching inside a separate method
        // to avoid the stack trace being populated from the test runner
        static Exception CreateMinimalException()
        {
            try
            {
                throw new InvalidOperationException("test");
            }
            catch (InvalidOperationException ex)
            {
                return ex;
            }
        }
    }

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
