using Serilog;
using Serilog.Events;

namespace SPUtil.App
{
    /// <summary>
    /// Thin wrapper around Serilog — creates a named logger for each class.
    /// Usage: private static readonly ILogger _log = AppLogger.For<MyClass>();
    /// </summary>
    public static class AppLogger
    {
        /// <summary>Returns a logger tagged with the class name as SourceContext.</summary>
        public static ILogger For<T>() =>
            Log.ForContext<T>();

        /// <summary>Returns a logger tagged with the given source context name.</summary>
        public static ILogger For(string context) =>
            Log.ForContext("SourceContext", context);
    }
}
