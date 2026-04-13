using System.Net.Http;

namespace BatchConvertToRVZ.services;

/// <summary>
/// Provides a shared HttpClientHandler instance for all HTTP services.
/// This ensures a single connection pool is used across the application.
/// </summary>
public static class SharedHttpHandler
{
    // Shared static HttpClient handler to prevent socket exhaustion.
    // Properly disposed when the application exits.
    private static readonly Lazy<SocketsHttpHandler> Handler = new(static () => new SocketsHttpHandler
    {
        PooledConnectionLifetime = TimeSpan.FromMinutes(2)
    });

    /// <summary>
    /// Gets the shared SocketsHttpHandler instance.
    /// </summary>
    public static SocketsHttpHandler Instance => Handler.Value;

    /// <summary>
    /// Disposes the shared handler when the application exits.
    /// Call this method during application shutdown.
    /// </summary>
    public static void Dispose()
    {
        if (Handler.IsValueCreated)
        {
            Handler.Value.Dispose();
        }
    }
}
