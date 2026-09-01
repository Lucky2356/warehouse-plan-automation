using System.Windows;
using System.Windows.Threading;
using WarehousePlanAutomation.Core.Logging;

namespace WarehousePlanAutomation.App;

public partial class App : Application
{
    private FileAppLogger? _logger;

    internal IAppLogger Logger => _logger ??= new FileAppLogger();

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;
        Logger.Information("Приложение запущено.");
    }

    protected override void OnExit(ExitEventArgs e)
    {
        Logger.Information("Приложение завершено.");
        base.OnExit(e);
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
