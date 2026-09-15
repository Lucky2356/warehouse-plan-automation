using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;

namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Запуск скрытого экземпляра Microsoft Excel и гарантированное его закрытие.
/// После Quit процесс проверяется по идентификатору окна: зависший EXCEL.EXE не остаётся.
/// </summary>
internal sealed class ExcelApplicationHost : IDisposable
{
    private readonly IAppLogger _logger;
    private readonly int _processId;
    private dynamic? _application;
    private bool _disposed;

    private ExcelApplicationHost(IAppLogger logger, dynamic application, int processId)
    {
        _logger = logger;
        _application = application;
        _processId = processId;
    }

    public dynamic Application => _application ?? throw new ObjectDisposedException(nameof(ExcelApplicationHost));

    public static ExcelApplicationHost Start(IAppLogger logger)
    {
        var progIdType = Type.GetTypeFromProgID("Excel.Application");
        if (progIdType is null)
        {
            throw new WarehousePlanException(
                "На этом компьютере не найден Microsoft Excel. " +
                "Программа работает через установленный Excel, поэтому его нужно установить и запустить хотя бы один раз.");
        }

        object? instance;
        try
        {
            instance = Activator.CreateInstance(progIdType);
        }
        catch (COMException ex)
        {
            throw new WarehousePlanException(
                "Не удалось запустить Microsoft Excel через COM. Закройте открытые окна Excel и повторите попытку.", ex);
        }

        if (instance is null)
        {
            throw new WarehousePlanException("Не удалось создать экземпляр Microsoft Excel.");
        }

        dynamic application = instance;
        var processId = ResolveProcessId(instance);

        application.Visible = false;
        application.DisplayAlerts = false;
        application.ScreenUpdating = false;
        application.EnableEvents = false;
        application.AskToUpdateLinks = false;

        logger.Information("Запущен Microsoft Excel, идентификатор процесса: " + processId);
        return new ExcelApplicationHost(logger, application, processId);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        var application = _application;
        _application = null;

        if (application is not null)
        {
            try
            {
                application.ScreenUpdating = true;
                application.EnableEvents = true;
                application.DisplayAlerts = true;
                application.Quit();
            }
            catch (COMException ex)
            {
                _logger.Warning("Excel не ответил на команду завершения.", ex);
            }
            finally
            {
                // Приложение больше нигде не используется: снимаем обёртку целиком,
                // чтобы ни одна живая ссылка не удерживала процесс EXCEL.EXE.
                ComUtils.FinalRelease(application);
            }
        }

        ComUtils.CollectGarbage();
        EnsureProcessExited();
    }

    private void EnsureProcessExited()
    {
        if (_processId <= 0)
        {
            return;
        }

        Process process;
        try
        {
            process = Process.GetProcessById(_processId);
        }
        catch (ArgumentException)
        {
            return;
        }

        using (process)
        {
            if (process.WaitForExit(5000))
            {
                return;
            }

            _logger.Warning("Процесс EXCEL.EXE " + _processId + " не завершился самостоятельно и будет закрыт принудительно.");
            try
            {
                process.Kill();
                process.WaitForExit(5000);
            }
            catch (InvalidOperationException)
            {
                // Процесс успел завершиться между проверкой и вызовом Kill.
            }
            catch (Win32Exception ex)
            {
                _logger.Warning("Не удалось завершить процесс EXCEL.EXE " + _processId + ".", ex);
            }
        }
    }

    private static int ResolveProcessId(object applicationObject)
    {
        dynamic application = applicationObject;
        try
        {
            int handle = application.Hwnd;
            _ = NativeMethods.GetWindowThreadProcessId(new IntPtr(handle), out var processId);
            return processId;
        }
        catch (COMException)
        {
            return 0;
        }
        catch (RuntimeBinderException)
        {
            return 0;
        }
    }

    private static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr windowHandle, out int processId);
    }
}
