using System.Globalization;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;

namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Путь к файлу результата. Исходный файл не изменяется никогда, поэтому результат
/// кладётся рядом отдельной книгой. Само имя собирает <see cref="OutputFileName"/>.
/// </summary>
internal static class OutputPath
{
    public static string Build(string sourceFilePath, DateTime now, string mark)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(sourceFilePath));
        if (string.IsNullOrEmpty(directory))
        {
            throw new WarehousePlanException("Не удалось определить папку исходного файла.");
        }

        var name = OutputFileName.Build(Path.GetFileNameWithoutExtension(sourceFilePath), mark, now);
        var extension = Path.GetExtension(sourceFilePath);

        var candidate = Path.Combine(directory, name + extension);
        var attempt = 1;
        while (File.Exists(candidate))
        {
            candidate = Path.Combine(
                directory, name + "_" + attempt.ToString(CultureInfo.InvariantCulture) + extension);
            attempt++;
        }

        return candidate;
    }

    public static void TryDelete(string path, IAppLogger logger)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException ex)
        {
            logger.Warning("Не удалось удалить частичный результат " + path + ".", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            logger.Warning("Нет прав на удаление частичного результата " + path + ".", ex);
        }
    }
}
