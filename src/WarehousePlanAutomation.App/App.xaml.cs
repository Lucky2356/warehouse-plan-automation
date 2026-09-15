using System.Windows;
using System.Windows.Threading;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Logging;

namespace WarehousePlanAutomation.App;

public partial class App : Application
{
    private FileAppLogger? _logger;
    private ThemeManager? _theme;
    private DispatcherTimer? _logCleanupTimer;

    internal IAppLogger Logger => FileLogger;

    private FileAppLogger FileLogger => _logger ??= new FileAppLogger();

    internal ThemeManager Theme => _theme ??= new ThemeManager(Logger);

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;

        // Тема ставится до создания окна: иначе оно на мгновение мелькнёт светлым.
        Theme.ApplySaved();

        Logger.Information("Приложение запущено.");
        ScheduleLogCleanup();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        Logger.Information("Приложение завершено.");
        _logCleanupTimer?.Stop();
        FileLogger.RemoveExpired(DateTime.Now);
        base.OnExit(e);
    }

    /// <summary>
    /// Вечерняя очистка журнала, пока программа открыта. Раз в минуту сверяется с часами:
    /// наступили ближайшие 20:00 - журнал чистится, и срок переносится на следующий вечер.
    /// Сверка по часам, а не отсчёт «через сутки»: компьютер мог спать, и проснувшийся
    /// сразу догоняет пропущенную очистку.
    /// </summary>
    private void ScheduleLogCleanup()
    {
        _logCleanupTimer = new DispatcherTimer(DispatcherPriority.Background) { Interval = TimeSpan.FromMinutes(1) };
        var next = LogCleanup.NextBoundary(DateTime.Now);
        _logCleanupTimer.Tick += (_, _) =>
        {
            var now = DateTime.Now;
            if (now < next)
            {
                return;
            }

            FileLogger.RemoveExpired(now);
            next = LogCleanup.NextBoundary(now);
        };
        _logCleanupTimer.Start();
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        Logger.Error("Необработанная ошибка интерфейса.", e.Exception);
        MessageBox.Show(
            "Произошла непредвиденная ошибка. Подробности записаны в журнал приложения." + Environment.NewLine +
            FileAppLogger.DefaultDirectory(),
            "Формирование плана склада",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        e.Handled = true;
    }

    private void OnDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
        {
            Logger.Error("Необработанная ошибка приложения.", exception);
        }
    }
}
