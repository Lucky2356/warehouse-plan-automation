using System.Diagnostics;
using System.Reflection;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Updates;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>
/// Обновления. Ручной работы здесь по замыслу нет: программа сама спрашивает GitHub
/// при запуске, сама скачивает новый файл и оставляет пользователю одну кнопку -
/// «Установить и перезапустить». Перезапуск не делается сам собой намеренно: в этот
/// момент может идти обработка книги, и закрывать программу без спроса нельзя.
/// </summary>
public sealed class UpdatesViewModel : NavPageViewModel
{
    private readonly IUpdateSource _source;
    private readonly IAppLogger _logger;
    private readonly Func<bool> _isBusyElsewhere;
    private readonly Action _shutdown;

    private AppRelease? _release;
    private string _downloadedPath = string.Empty;

    private string _statusMessage = "Проверка обновлений...";
    private string _details = string.Empty;
    private string _notes = string.Empty;
    private bool _isChecking;
    private bool _isDownloading;
    private bool _isReady;
    private bool _hasProblem;
    private int _progressValue;

    public UpdatesViewModel(
        IUpdateSource source,
        IAppLogger logger,
        Func<bool> isBusyElsewhere,
        Action shutdown)
    {
        _source = source;
        _logger = logger;
        _isBusyElsewhere = isBusyElsewhere;
        _shutdown = shutdown;

        CheckCommand = new RelayCommand(() => _ = RunAsync(), () => !IsChecking && !IsDownloading);
        InstallCommand = new RelayCommand(Install, () => IsReady);
        OpenPageCommand = new RelayCommand(OpenPage);
    }

    public override string Title => "Обновления";

    public override string Subtitle => "Новые версии программы с GitHub";

    /// <summary>Стрелка вниз в лоток: скачать и поставить.</summary>
    public override string IconData =>
        "M10,3.4 L10,11.8 " +
        "M6.6,8.6 L10,12 L13.4,8.6 " +
        "M4,14.2 L4,15.6 A1.4,1.4 0 0,0 5.4,17 L14.6,17 A1.4,1.4 0 0,0 16,15.6 L16,14.2";

    /// <summary>Версия работающей программы.</summary>
    public string CurrentVersion { get; } = ReadCurrentVersion();

    public RelayCommand CheckCommand { get; }

    public RelayCommand InstallCommand { get; }

    public RelayCommand OpenPageCommand { get; }

    public string StatusMessage
    {
        get => _statusMessage;
        private set => SetProperty(ref _statusMessage, value);
    }

    /// <summary>Строка под заголовком: версия, дата, размер файла.</summary>
    public string Details
    {
        get => _details;
        private set
        {
            SetProperty(ref _details, value);
            OnPropertyChanged(nameof(HasDetails));
        }
    }

    public bool HasDetails => Details.Length > 0;

    /// <summary>Описание выпуска, как оно написано на GitHub.</summary>
    public string Notes
    {
        get => _notes;
        private set
        {
            SetProperty(ref _notes, value);
            OnPropertyChanged(nameof(HasNotes));
        }
    }

    public bool HasNotes => Notes.Length > 0;

    public bool IsChecking
    {
        get => _isChecking;
        private set
        {
            SetProperty(ref _isChecking, value);
            RefreshCommands();
        }
    }

    public bool IsDownloading
    {
        get => _isDownloading;
        private set
        {
            SetProperty(ref _isDownloading, value);
            RefreshCommands();
        }
    }

    /// <summary>Файл скачан и лежит рядом: остался один щелчок.</summary>
    public bool IsReady
    {
        get => _isReady;
        private set
        {
            SetProperty(ref _isReady, value);
            RefreshCommands();
        }
    }

    public bool HasProblem
    {
        get => _hasProblem;
        private set => SetProperty(ref _hasProblem, value);
    }

    public int ProgressValue
    {
        get => _progressValue;
        private set => SetProperty(ref _progressValue, value);
    }

    /// <summary>Проверка при запуске окна. Ошибка сети здесь не должна ничему мешать.</summary>
    public void CheckOnStartup() => _ = RunAsync();

    private async Task RunAsync()
    {
        if (IsChecking || IsDownloading)
        {
            return;
        }

        IsChecking = true;
        IsReady = false;
        HasProblem = false;
        ProgressValue = 0;
        Details = string.Empty;
        Notes = string.Empty;
        StatusMessage = "Проверка обновлений...";

        UpdateCheck check;
        try
        {
            check = await _source.CheckAsync(CurrentVersionValue, CancellationToken.None)
                .ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            _logger.Error("Проверка обновлений сорвалась.", ex);
            Fail("Проверить обновления не удалось. Подробности в журнале.");
            IsChecking = false;
            return;
        }

        IsChecking = false;

        if (check.Problem is not null)
        {
            Fail(check.Problem);
            return;
        }

        if (!check.HasUpdate)
        {
            Badge = string.Empty;
            StatusMessage = "Установлена последняя версия.";
            Details = "Версия " + CurrentVersion + ".";
            return;
        }

        _release = check.Release;
        Badge = "новое";
        Notes = _release!.Notes;
        Details = Describe(_release);
        StatusMessage = "Есть новая версия " + AppVersion.Display(_release.Version) + ". Загрузка...";

        await DownloadAsync(_release).ConfigureAwait(true);
    }

    private async Task DownloadAsync(AppRelease release)
    {
        IsDownloading = true;
        ProgressValue = 0;

        try
        {
            var progress = new Progress<int>(percent => ProgressValue = percent);
            _downloadedPath = await _source
                .DownloadAsync(release, progress, CancellationToken.None)
                .ConfigureAwait(true);

            ProgressValue = 100;
            IsReady = true;
            StatusMessage =
                "Версия " + AppVersion.Display(release.Version) +
                " скачана. Нажмите «Установить и перезапустить».";
        }
        catch (UpdateSignatureException ex)
        {
            _logger.Error("Обновление не прошло проверку подлинности.", ex);
            Fail(
                "Новая версия " + AppVersion.Display(release.Version) +
                " не установлена: не прошла проверку подлинности. " + ex.Message);
        }
        catch (Exception ex)
        {
            _logger.Error("Не удалось скачать обновление.", ex);
            Fail(
                "Новая версия " + AppVersion.Display(release.Version) +
                " есть, но скачать её не удалось. Можно взять файл вручную на GitHub.");
        }
        finally
        {
            IsDownloading = false;
        }
    }

    /// <summary>
    /// Подменяет файл и закрывает программу. Если в этот момент идёт обработка книги,
    /// установка откладывается: прервать её на середине хуже, чем обновиться позже.
    /// </summary>
    private void Install()
    {
        if (_release is null || _downloadedPath.Length == 0)
        {
            return;
        }

        if (_isBusyElsewhere())
        {
            StatusMessage =
                "Сейчас идёт обработка книги. Дождитесь её окончания и нажмите ещё раз - " +
                "обновление уже скачано.";
            return;
        }

        // Между загрузкой и установкой файл лежит во временной папке: подпись проверяется
        // ещё раз прямо перед запуском, чтобы поставить ровно то, что проверили.
        try
        {
            var signature = System.IO.File.ReadAllText(_downloadedPath + UpdateSignature.Extension);
            UpdateSignature.EnsureValid(_downloadedPath, _release.AssetName, signature);
        }
        catch (Exception ex) when (ex is UpdateSignatureException or System.IO.IOException or UnauthorizedAccessException)
        {
            _logger.Error("Перед установкой подпись обновления не сошлась.", ex);
            Fail("Обновление не установлено: файл не прошёл проверку подлинности. Скачайте его заново.");
            IsReady = false;
            return;
        }

        if (!UpdateInstaller.Launch(_downloadedPath, _release, _logger))
        {
            Fail("Установить обновление не удалось. Файл лежит в " + _downloadedPath);
            return;
        }

        StatusMessage = "Программа закроется и откроется заново уже обновлённой.";
        _shutdown();
    }

    private void OpenPage()
    {
        try
        {
            Process.Start(new ProcessStartInfo(_source.ReleasesPage) { UseShellExecute = true });
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            _logger.Warning("Не удалось открыть страницу выпусков.", ex);
        }
    }

    private void Fail(string message)
    {
        Badge = string.Empty;
        HasProblem = true;
        IsReady = false;
        StatusMessage = message;
    }

    private static string Describe(AppRelease release)
    {
        var parts = new List<string> { "Версия " + AppVersion.Display(release.Version) };

        if (release.Published is not null)
        {
            parts.Add("от " + release.Published.Value.ToLocalTime().ToString("dd.MM.yyyy"));
        }

        if (release.AssetSize > 0)
        {
            parts.Add((release.AssetSize / 1024d / 1024d).ToString("0.#") + " МБ");
        }

        return string.Join(", ", parts) + ".";
    }

    private Version CurrentVersionValue { get; } =
        Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0, 0);

    private static string ReadCurrentVersion()
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        return version is null ? "неизвестна" : AppVersion.Display(version);
    }

    private void RefreshCommands()
    {
        CheckCommand.RaiseCanExecuteChanged();
        InstallCommand.RaiseCanExecuteChanged();
    }
}
