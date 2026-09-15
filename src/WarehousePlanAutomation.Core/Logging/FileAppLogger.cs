using System.Globalization;
using System.Text;

namespace WarehousePlanAutomation.Core.Logging;

/// <summary>
/// Технический лог в %LOCALAPPDATA%\WarehousePlanAutomation\Logs.
/// Пишутся этапы обработки, количества строк и полные исключения; данные Excel не выгружаются.
///
/// Лог хранится недолго: каждый вечер в <see cref="LogCleanup.EveningHour"/>:00 удаляется всё,
/// что записано до этого времени (см. <see cref="LogCleanup"/>).
/// </summary>
public sealed class FileAppLogger : IAppLogger
{
    private const string FilePattern = "warehouse-plan-*.log";

    private readonly object _sync = new();
    private readonly string _directory;

    public FileAppLogger(string? directory = null)
    {
        _directory = directory ?? DefaultDirectory();
        System.IO.Directory.CreateDirectory(_directory);
        RemoveExpired(DateTime.Now);
    }

    public string LogDirectory => _directory;

    public static string DefaultDirectory() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WarehousePlanAutomation",
        "Logs");

    public string CurrentFilePath =>
        Path.Combine(_directory, "warehouse-plan-" + DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + ".log");

    public void Log(LogLevel level, string message, Exception? exception = null)
    {
        var builder = new StringBuilder();
        builder.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
        builder.Append(" [").Append(level.ToString().ToUpperInvariant()).Append("] ");
        builder.Append(message);

        if (exception is not null)
        {
            builder.AppendLine();
            builder.Append(exception);
        }

        builder.AppendLine();

        lock (_sync)
        {
            try
            {
                File.AppendAllText(CurrentFilePath, builder.ToString(), Encoding.UTF8);
            }
            catch (IOException)
            {
                // Журналирование не должно прерывать обработку файла пользователя.
            }
            catch (UnauthorizedAccessException)
            {
                // То же самое: отсутствие прав на запись лога не является ошибкой обработки.
            }
        }
    }

    /// <summary>
    /// Удаляет файлы журнала, последняя запись в которых сделана до последнего вечернего срока.
    /// Вызывается при запуске, в сам вечерний срок, пока программа открыта, и при выходе.
    /// Файл пишется открытием на дописывание и сразу закрывается, поэтому удалению он не мешает.
    /// </summary>
    public int RemoveExpired(DateTime now)
    {
        var boundary = LogCleanup.LastBoundary(now);
        var removed = 0;

        lock (_sync)
        {
            try
            {
                foreach (var file in System.IO.Directory.GetFiles(_directory, FilePattern))
                {
                    if (File.GetLastWriteTime(file) < boundary)
                    {
                        File.Delete(file);
                        removed++;
                    }
                }
            }
            catch (IOException)
            {
                // Очистка журнала не должна мешать работе; попробуем в следующий раз.
            }
            catch (UnauthorizedAccessException)
            {
                // То же самое.
            }
        }

        return removed;
    }
}

/// <summary>
/// Когда удаляется журнал. Каждый вечер в 20:00 стирается всё, что записано до этого времени.
/// Программа не служба и может быть закрыта в 20:00, поэтому срок проверяется в трёх местах:
/// при запуске, в 20:00, если программа открыта, и при закрытии. Запись, сделанная после 20:00,
/// живёт до следующего вечера.
/// </summary>
public static class LogCleanup
{
    public const int EveningHour = 20;

    /// <summary>Последний прошедший вечерний срок: сегодня в 20:00, а до 20:00 - вчера в 20:00.</summary>
    public static DateTime LastBoundary(DateTime now)
    {
        var today = now.Date.AddHours(EveningHour);
        return now >= today ? today : today.AddDays(-1);
    }

    /// <summary>Следующий вечерний срок - когда таймер программы должен сработать.</summary>
    public static DateTime NextBoundary(DateTime now) => LastBoundary(now).AddDays(1);
}
