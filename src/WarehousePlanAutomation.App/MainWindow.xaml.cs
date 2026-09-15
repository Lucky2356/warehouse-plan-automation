using System.Windows;
using System.Windows.Media;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.App.ViewModels;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Updates;
using WarehousePlanAutomation.Excel;

namespace WarehousePlanAutomation.App;

public partial class MainWindow : Window
{
    /// <summary>Репозиторий, откуда программа берёт свои же выпуски.</summary>
    private const string RepositoryOwner = "Lucky2356";

    private const string RepositoryName = "warehouse-plan-automation";

    private readonly ShellViewModel _shell;
    private readonly GitHubUpdateSource _updates;

    public MainWindow()
    {
        InitializeComponent();

        var logger = ((App)Application.Current).Logger;
        var prompt = new DialogDecisionPrompt(this);
        var settings = new UserSettings(logger);

        _updates = new GitHubUpdateSource(logger, RepositoryOwner, RepositoryName);

        var tasks = new WorkbookTaskViewModel[]
        {
            new PlanTaskViewModel(new ExcelWorkbookProcessor(logger), logger, settings),
            new PriceTaskViewModel(
                new ExcelPriceSheetProcessor(logger, prompt, PriceStage.Prepare), logger, settings),
            new DistributionTaskViewModel(
                new ExcelPriceSheetProcessor(
                    logger,
                    prompt,
                    PriceStage.Recalculate,
                    approvedPricesFolder: () => settings.Get(DistributionTaskViewModel.ApprovedPricesFolderKey)),
                logger,
                settings),
            new RestockTaskViewModel(new ExcelRestockProcessor(logger), logger, settings),
            new ReceivingPrepareTaskViewModel(
                new ExcelReceivingProcessor(logger, ReceivingStage.Prepare), logger, settings),
            new ReceivingAddressesTaskViewModel(
                new ExcelReceivingProcessor(logger, ReceivingStage.Addresses), logger, settings),
        };

        var updatesPage = new UpdatesViewModel(
            _updates,
            logger,
            () => tasks.Any(task => task.IsBusy),
            () => Application.Current.Shutdown());

        _shell = new ShellViewModel(
            tasks.Cast<NavPageViewModel>().Append(updatesPage).ToArray(),
            ((App)Application.Current).Theme);

        DataContext = _shell;

        // Заголовок окна перекрашивается вместе с темой: сам он за ней не следует.
        _shell.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(ShellViewModel.IsDarkTheme))
            {
                ApplyTitleBar();
            }
        };

        // Проверка идёт сразу при открытии окна и ничего не блокирует: если сети нет,
        // на вкладке обновлений просто будет написано, что связаться не удалось.
        Loaded += (_, _) => updatesPage.CheckOnStartup();
    }

    /// <summary>Ниже этой ширины боковая панель сворачивается в значки.</summary>
    private const double CompactWidth = 1120;

    /// <summary>Задача, к которой относятся выбор файла и перетаскивание.</summary>
    private WorkbookTaskViewModel? Current => _shell.SelectedTask;

    /// <summary>
    /// Размер окна берётся от рабочей области экрана, а не задаётся числом: на ноутбуке
    /// 1366×768 фиксированные 1180×760 не помещаются, на большом мониторе выглядят
    /// форточкой. Доля подобрана так, чтобы рядом оставалось место для окна Excel.
    /// </summary>
    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);

        var work = SystemParameters.WorkArea;
        Width = Math.Round(Math.Clamp(work.Width * 0.68, MinWidth, 1320));
        Height = Math.Round(Math.Clamp(work.Height * 0.84, MinHeight, 900));

        Left = work.Left + ((work.Width - Width) / 2);
        Top = work.Top + ((work.Height - Height) / 2);

        ApplyCompact();
        ApplyTitleBar();
    }

    /// <summary>
    /// Заголовок окна рисует Windows, и сам он под тему не подстраивается - его цвет
    /// приходится задавать отдельно, при запуске и при каждом переключении темы.
    /// </summary>
    private void ApplyTitleBar() => TitleBarTheme.ApplyCurrent(this);

    private void OnSizeChanged(object sender, SizeChangedEventArgs e) => ApplyCompact();

    private void ApplyCompact() =>
        _shell.IsCompact = ActualWidth > 0 && ActualWidth < CompactWidth;

    protected override void OnClosed(EventArgs e)
    {
        foreach (var task in _shell.Tasks)
        {
            task.Cancel();
        }

        _updates.Dispose();
        base.OnClosed(e);
    }

    private void OnDragOver(object sender, DragEventArgs e)
    {
        var task = Current;
        var accepted = task is { IsBusy: false } && TryGetWorkbookPath(e, out _);

        e.Effects = accepted ? DragDropEffects.Copy : DragDropEffects.None;
        if (task is not null)
        {
            task.IsDragOver = accepted;
        }

        e.Handled = true;
    }

    private void OnDragLeave(object sender, DragEventArgs e)
    {
        if (Current is { } task)
        {
            task.IsDragOver = false;
        }

        e.Handled = true;
    }

    private void OnDrop(object sender, DragEventArgs e)
    {
        e.Handled = true;

        if (Current is not { } task)
        {
            return;
        }

        task.IsDragOver = false;

        if (TryGetWorkbookPath(e, out var path))
        {
            task.ApplySelectedFile(path);
        }
    }

    /// <summary>
    /// Из перетаскиваемого набора берётся первый файл: книга у обработки всегда одна.
    /// Формат проверяется тем же списком расширений, что и в диалоге выбора.
    /// </summary>
    private static bool TryGetWorkbookPath(DragEventArgs e, out string path)
    {
        path = string.Empty;

        if (!e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            return false;
        }

        if (e.Data.GetData(DataFormats.FileDrop) is not string[] { Length: > 0 } files)
        {
            return false;
        }

        if (!WorkbookFile.IsSupported(files[0]))
        {
            return false;
        }

        path = files[0];
        return true;
    }
}
