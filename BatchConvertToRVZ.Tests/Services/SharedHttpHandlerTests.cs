using BatchConvertToRVZ.services;
using Xunit;

namespace BatchConvertToRVZ.Tests.Services;

public class SharedHttpHandlerTests
{
    [Fact]
    public void InstanceReturnsSameObjectOnMultipleCalls()
    {
        var instance1 = SharedHttpHandler.Instance;
        var instance2 = SharedHttpHandler.Instance;

        Assert.Same(instance1, instance2);
    }

    [Fact]
    public void InstanceIsOfTypeSocketsHttpHandler()
    {
        var instance = SharedHttpHandler.Instance;

        Assert.IsType<SocketsHttpHandler>(instance);
    }

    [Fact]
    public void PooledConnectionLifetimeIsTwoMinutes()
    {
        var instance = SharedHttpHandler.Instance;

        Assert.Equal(TimeSpan.FromMinutes(2), instance.PooledConnectionLifetime);
    }

    [Fact]
    public void DisposeCanBeCalledMultipleTimesWithoutException()
    {
        // First call may or may not dispose depending on if Instance was already accessed
        SharedHttpHandler.Dispose();

        // Second call should not throw
        var exception = Record.Exception(static () => SharedHttpHandler.Dispose());

        Assert.Null(exception);
    }
}
