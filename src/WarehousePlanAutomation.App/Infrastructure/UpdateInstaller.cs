using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Updates;

namespace WarehousePlanAutomation.App.Infrastructure;

/// <summary>
/// Установка обновления: программа не может переписать сама себя, пока работает,
/// поэтому подмену делает маленький сценарий cmd. Он ждёт, пока процесс закроется,
/// ставит новый файл на место старого и запускает его.
///
/// Все пути передаются сценарию переменными окружения, а не текстом внутри файла:
/// сам файл тогда остаётся набором латинских букв и не зависит от кодовой страницы
/// консоли, даже если папка пользователя названа по-русски.
/// </summary>
public static class UpdateInstaller
{
    private const string Script = """
        @echo off
        :wait
        tasklist /FI "PID eq %UPD_PID%" /NH 2>nul | find /I "%UPD_EXE%" >nul
        if not errorlevel 1 (
          ping -n 2 127.0.0.1 >nul
          goto wait
        )
        set /a UPD_TRY=0
        :retry
        move /y "%UPD_SRC%" "%UPD_DST%" >nul 2>&1
        if not errorlevel 1 goto done
        set /a UPD_TRY+=1
        if %UPD_TRY% GEQ 30 goto fail
        ping -n 2 127.0.0.1 >nul
        goto retry
        :done
        if /i not "%UPD_DST%"=="%UPD_OLD%" del /f /q "%UPD_OLD%" >nul 2>&1
        start "" "%UPD_DST%"
        goto end
        :fail
        start "" "%UPD_SRC%"
        :end
        (goto) 2>nul & del /f /q "%~f0"
        """;

    /// <summary>
    /// Запускает подмену и сообщает, удалось ли её начать. Программу после этого
    /// нужно закрыть: пока она работает, файл занят и сценарий будет ждать.
    /// </summary>
    public static bool Launch(string downloadedPath, AppRelease release, IAppLogger logger)
    {
        var current = Environment.ProcessPath;
        if (string.IsNullOrEmpty(current))
        {
            logger.Error("Не удалось определить путь к собственному файлу: обновление не установлено.");
            return false;
        }

        if (!File.Exists(downloadedPath))
        {
            logger.Error("Файл обновления не найден: " + downloadedPath);
            return false;
        }

        var target = BuildTargetPath(current, release);

        try
        {
            var script = Path.Combine(
                Path.GetDirectoryName(downloadedPath) ?? Path.GetTempPath(), "install.cmd");
            File.WriteAllText(script, Script, new UTF8Encoding(false));

            var start = new ProcessStartInfo("cmd.exe", "/c \"" + script + "\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetTempPath(),
            };

            start.Environment["UPD_PID"] = Environment.ProcessId.ToString(CultureInfo.InvariantCulture);
            start.Environment["UPD_EXE"] = Path.GetFileName(current);
            start.Environment["UPD_SRC"] = downloadedPath;
            start.Environment["UPD_DST"] = target;
            start.Environment["UPD_OLD"] = current;

            Process.Start(start);

            logger.Information(
                "Запущена установка обновления: " + downloadedPath + " -> " + target +
                (string.Equals(target, current, StringComparison.OrdinalIgnoreCase)
                    ? " (на место текущего файла)"
                    : ", прежний файл " + current + " будет удалён"));
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or System.ComponentModel.Win32Exception)
        {
            logger.Error("Не удалось запустить установку обновления.", ex);
            return false;
        }
    }

    /// <summary>
    /// Куда положить новый файл. Если в имени текущего стоит его версия
    /// («WarehousePlanAutomation-1.8.0-win-x64.exe»), новый файл получает своё имя
    /// из выпуска и кладётся рядом, а прежний удаляется - иначе имя врало бы о версии.
    /// В остальных случаях файл заменяется на месте: путь мог попасть в ярлык.
    /// </summary>
    internal static string BuildTargetPath(string currentPath, AppRelease release)
    {
        var directory = Path.GetDirectoryName(currentPath);
        var name = Path.GetFileName(currentPath);

        if (directory is null || release.AssetName.Length == 0)
        {
            return currentPath;
        }

        var version = typeof(UpdateInstaller).Assembly.GetName().Version;
        var marker = version is null ? null : AppVersion.Display(version);

        return marker is not null && name.Contains(marker, StringComparison.Ordinal)
            ? Path.Combine(directory, release.AssetName)
            : currentPath;
    }
}
