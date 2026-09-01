using System.Globalization;
using System.Text;

namespace WarehousePlanAutomation.Core.Logging;

/// <summary>
/// Технический лог в %LOCALAPPDATA%\WarehousePlanAutomation\Logs.
/// Пишутся этапы обработки, количества строк и полные исключения; данные Excel не выгружаются.
/// </summary>
public sealed class FileAppLogger : IAppLogger
{
    private readonly object _sync = new();
    private readonly string _directory;
    private readonly int _retainDays;

    public FileAppLogger(string? directory = null, int retainDays = 30)
    {
        _directory = directory ?? DefaultDirectory();
        _retainDays = retainDays;
        System.IO.Directory.CreateDirectory(_directory);
        RemoveOldFiles();
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

    private void RemoveOldFiles()
    {
        try
        {
            var threshold = DateTime.Now.AddDays(-_retainDays);
            foreach (var file in System.IO.Directory.GetFiles(_directory, "warehouse-plan-*.log"))
            {
                if (File.GetLastWriteTime(file) < threshold)
                {
                    File.Delete(file);
                }
            }
        }
        catch (IOException)
        {
            // Очистка старых логов не критична.
        }
        catch (UnauthorizedAccessException)
        {
            // Очистка старых логов не критична.
        }
    }
}
