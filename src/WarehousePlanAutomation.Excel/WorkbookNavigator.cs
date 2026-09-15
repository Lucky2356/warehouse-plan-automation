using System.Runtime.InteropServices;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;

namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Открывает готовую книгу на нужном месте: лист, ячейка, курсор уже там.
///
/// Главное здесь - не открыть вторую копию. Если книга уже открыта у аналитика,
/// новый экземпляр Excel открыл бы её только для чтения, а несохранённые правки
/// остались бы в первом. Поэтому программа сначала пытается присоединиться
/// к работающему Excel и только потом запускает свой.
/// </summary>
public static class WorkbookNavigator
{
    private const string ProgId = "Excel.Application";

    /// <summary>Excel уже открыт этим пользователем - в таблице запущенных объектов.</summary>
    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(
        ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object instance);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr window);

    /// <summary>
    /// Показывает <paramref name="cell"/> листа <paramref name="sheet"/> книги
    /// <paramref name="path"/>. Пустая ячейка - открывается просто лист.
    /// </summary>
    public static Task ShowAsync(
        string path, string? sheet, string? cell, IAppLogger logger) =>
        StaTaskRunner.RunAsync(() => Show(path, sheet, cell, logger), CancellationToken.None);

    private static bool Show(string path, string? sheet, string? cell, IAppLogger logger)
    {
        if (!File.Exists(path))
        {
            throw new WarehousePlanException("Файл результата не найден: " + path);
        }

        // Все обёртки освобождаются перед выходом. Без этого живая ссылка на Excel
        // осталась бы в программе, и EXCEL.EXE продолжал бы висеть в памяти уже после
        // того, как аналитик закрыл окно книги.
        using var scope = new ComScope();
        object applicationObject = Attach(logger) ?? Create();
        dynamic application = applicationObject;

        try
        {
            application.Visible = true;

            dynamic workbooks = scope.Track(application.Workbooks);
            dynamic workbook = scope.Track(FindOpen(workbooks, path) ?? workbooks.Open(path));
            workbook.Activate();

            if (!string.IsNullOrEmpty(sheet))
            {
                dynamic sheets = scope.Track(workbook.Worksheets);
                dynamic worksheet = scope.Track(sheets[sheet]);
                worksheet.Activate();

                if (!string.IsNullOrEmpty(cell))
                {
                    dynamic range = scope.Track(worksheet.Range[cell]);

                    // Прокрутка отдельно от выделения: Select ставит курсор, но если
                    // ячейка далеко внизу, окно так и останется на первой строке.
                    application.Goto(range, true);
                    range.Select();
                }
            }

            Focus(application, logger);
            logger.Information(
                "Открыта книга " + path +
                (string.IsNullOrEmpty(sheet) ? string.Empty : ", лист «" + sheet + "»") +
                (string.IsNullOrEmpty(cell) ? string.Empty : ", ячейка " + cell) + ".");
            return true;
        }
        catch (COMException ex)
        {
            logger.Error("Не удалось открыть книгу на нужном месте.", ex);
            throw new WarehousePlanException(
                "Не удалось открыть книгу на нужном месте. Откройте файл вручную.", ex);
        }
        finally
        {
            // Приложение освобождается полностью: дальше им распоряжается человек,
            // а у программы на него ссылок остаться не должно.
            ComUtils.FinalRelease(applicationObject);
        }
    }

    /// <summary>Книга уже открыта в этом Excel - тогда вторую копию открывать нельзя.</summary>
    private static object? FindOpen(dynamic workbooks, string path)
    {
        var full = Path.GetFullPath(path);

        foreach (dynamic workbook in workbooks)
        {
            string opened;
            try
            {
                opened = workbook.FullName;
            }
            catch (COMException)
            {
                continue;
            }

            if (string.Equals(opened, full, StringComparison.OrdinalIgnoreCase))
            {
                return workbook;
            }

            ComUtils.Release(workbook);
        }

        return null;
    }

    private static object? Attach(IAppLogger logger)
    {
        try
        {
            if (CLSIDFromProgID(ProgId, out var clsid) != 0)
            {
                return null;
            }

            return GetActiveObject(ref clsid, IntPtr.Zero, out var instance) == 0 ? instance : null;
        }
        catch (COMException ex)
        {
            logger.Debug("Работающий Excel не найден: " + ex.Message);
            return null;
        }
    }

    private static object Create()
    {
        var type = Type.GetTypeFromProgID(ProgId)
                   ?? throw new WarehousePlanException(
                       "Не удалось запустить Microsoft Excel: программа не установлена.");

        return Activator.CreateInstance(type)
               ?? throw new WarehousePlanException("Не удалось запустить Microsoft Excel.");
    }

    /// <summary>Excel уже открыт, но за другими окнами - без этого перехода не видно.</summary>
    private static void Focus(dynamic application, IAppLogger logger)
    {
        try
        {
            var handle = new IntPtr((long)application.Hwnd);
            if (handle != IntPtr.Zero)
            {
                SetForegroundWindow(handle);
            }
        }
        catch (COMException ex)
        {
            logger.Debug("Не удалось вывести окно Excel вперёд: " + ex.Message);
        }
    }
}
