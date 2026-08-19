using System.IO;
using System.Runtime.CompilerServices;

using Serilog;

namespace SPUtil.Infrastructure
{
    /// <summary>
    /// Attaches the caller's method name and source file to a log entry as a
    /// "Caller" property, without touching Serilog's existing message-template
    /// / params call syntax.
    ///
    /// Usage:
    ///     _log.Caller().Error(ex, "AddWebPartAsync failed for {Page}", pageUrl);
    ///     _log.Caller().Warning("Placeholder not found for '{Title}'", title);
    ///
    /// [CallerMemberName] / [CallerFilePath] are filled in by the compiler at
    /// the call site as long as Caller() is invoked with no explicit arguments.
    /// Do NOT pass explicit values for callerMember/callerFile at call sites —
    /// doing so defeats the purpose and reports the wrong location.
    /// </summary>
    public static class LoggerCallerExtensions
    {
        public static ILogger Caller(
            this ILogger logger,
            [CallerMemberName] string callerMember = "",
            [CallerFilePath] string callerFile = "")
        {
            string fileName = Path.GetFileNameWithoutExtension(callerFile);
            return logger.ForContext("Caller", $"{fileName}.{callerMember}");
        }
    }
}
