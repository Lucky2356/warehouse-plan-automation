namespace WarehousePlanAutomation.Core.Logging;

public enum LogLevel
{
    Debug,
    Information,
    Warning,
    Error,
}

/// <summary>Минимальная абстракция журналирования, не тянущая внешних зависимостей.</summary>
public interface IAppLogger
{
    void Log(LogLevel level, string message, Exception? exception = null);
}

public static class AppLoggerExtensions
{
    public static void Debug(this IAppLogger logger, string message) =>
        logger.Log(LogLevel.Debug, message);

    public static void Information(this IAppLogger logger, string message) =>
        logger.Log(LogLevel.Information, message);

    public static void Warning(this IAppLogger logger, string message, Exception? exception = null) =>
        logger.Log(LogLevel.Warning, message, exception);

    public static void Error(this IAppLogger logger, string message, Exception? exception = null) =>
        logger.Log(LogLevel.Error, message, exception);
}

/// <summary>Логгер, ничего не пишущий. Используется в тестах.</summary>
public sealed class NullAppLogger : IAppLogger
{
    public static readonly NullAppLogger Instance = new();

    public void Log(LogLevel level, string message, Exception? exception = null)
    {
    }
}
