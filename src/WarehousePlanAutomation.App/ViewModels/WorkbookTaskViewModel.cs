using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Excel;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>
/// Задача «одна книга на входе - одна книга на выходе». Всё, что одинаково у любой
/// такой задачи, живёт здесь: выбор файла, перетаскивание, прогресс, отмена, карточка
/// результата. Наследник задаёт только тексты и свой обработчик.
/// </summary>
public abstract class WorkbookTaskViewModel : NavPageViewModel
{
    /// <summary>Ключ, под которым задача помнит свою папку.</summary>
    private const string FolderKey = "lastFolder.";

    private readonly IWorkbookProcessor _processor;
    private readonly IAppLogger _logger;
    private readonly UserSettings? _settings;

    private string _selectedFilePath = string.Empty;
    private string _statusMessage;
    private string _errorMessage = string.Empty;
    private string _resultPath = string.Empty;
    private bool _isBusy;
    private bool _isDragOver;
    private int _progressValue;
    private bool _isNavigating;
    private bool _preflightDone;
    private string _optionFolderPath;
    private string? _selectedRunOption;

    private CancellationTokenSource? _cancellation;

    protected WorkbookTaskViewModel(
        IWorkbookProcessor processor, IAppLogger logger, UserSettings? settings = null)
    {
        _processor = processor;
        _logger = logger;
        _settings = settings;
        _statusMessage = InitialHint;
        _optionFolderPath = FolderOptionKey is null ? string.Empty : settings?.Get(FolderOptionKey) ?? string.Empty;

        SelectOptionFolderCommand = new RelayCommand(SelectOptionFolder, () => !IsBusy && HasFolderOption);
        SelectOptionFileCommand = new RelayCommand(SelectOptionFile, () => !IsBusy && FolderOptionAllowsFile);
        ClearOptionFolderCommand = new RelayCommand(
            () => OptionFolderPath = string.Empty,
            () => !IsBusy && HasOptionFolder);

        SelectFileCommand = new RelayCommand(SelectFile, () => !IsBusy);
        ProcessCommand = new RelayCommand(
            () => _ = RunAsync(),
            () => !IsBusy && SelectedFilePath.Length > 0);
        CancelCommand = new RelayCommand(Cancel, () => IsBusy);
        OpenResultFolderCommand = new RelayCommand(OpenResultFolder, () => ResultPath.Length > 0);
        OpenInExcelCommand = new RelayCommand(
            () => _ = OpenInExcelAsync(),
            () => ResultPath.Length > 0 && !_isNavigating);
        OpenWarningCommand = new RelayCommand<ProcessingWarning>(
            warning => _ = OpenWarningAsync(warning),
            warning => warning.CanNavigate && ResultPath.Length > 0);
    }

    /// <summary>Что написано на кнопке запуска.</summary>
    public abstract string ActionCaption { get; }

    /// <summary>Подсказка до выбора файла.</summary>
    public abstract string InitialHint { get; }

    /// <summary>Сообщение об успехе.</summary>
    public abstract string SuccessMessage { get; }

    /// <summary>Начало сообщения об ошибке: «Не удалось ...».</summary>
    public abstract string FailureMessage { get; }

    /// <summary>Короткое имя задачи для настроек - им подписана запомненная папка.</summary>
    protected abstract string SettingsKey { get; }

    /// <summary>
    /// Листы, без которых задача не выполнится. По ним книга проверяется сразу после
    /// выбора файла - до того, как человек нажмёт кнопку и будет ждать полминуты.
    /// </summary>
    protected abstract IReadOnlyList<string> RequiredSheets { get; }

    /// <summary>
    /// Листы, без которых задача сделает не всё. Их нехватка - предупреждение, а не помеха:
    /// работа запустится и сделает ту часть, для которой данных хватает.
    /// </summary>
    protected virtual IReadOnlyDictionary<string, string> OptionalSheets { get; } =
        new Dictionary<string, string>();

    /// <summary>
    /// Обязательные листы, которых в книге может быть несколько - обработка берёт их все.
    /// Для них несколько подходящих листов не помеха.
    /// </summary>
    protected virtual IReadOnlyCollection<string> RepeatableSheets { get; } = Array.Empty<string>();

    /// <summary>
    /// Ключ настройки дополнительной папки, которую задача читает при запуске. null - папки
    /// у задачи нет, и в окне её поле не показывается.
    /// </summary>
    protected virtual string? FolderOptionKey => null;

    /// <summary>Подпись поля дополнительной папки.</summary>
    public virtual string FolderOptionTitle => string.Empty;

    /// <summary>Что будет, если папку выбрать, и что - если нет.</summary>
    public virtual string FolderOptionHint => string.Empty;

    public bool HasFolderOption => FolderOptionKey is not null;

    /// <summary>
    /// Вместо папки можно выбрать и файл: в окне выбора папки файлы не видны, и человеку,
    /// который знает нужный файл, приходится искать его вслепую.
    /// </summary>
    public virtual bool FolderOptionAllowsFile => false;

    /// <summary>Выбранная дополнительная папка. Запоминается сразу: обработчик читает её при запуске.</summary>
    public string OptionFolderPath
    {
        get => _optionFolderPath;
        private set
        {
            SetProperty(ref _optionFolderPath, value);
            OnPropertyChanged(nameof(HasOptionFolder));
            if (FolderOptionKey is not null)
            {
                _settings?.Set(FolderOptionKey, value);
            }

            RefreshCommands();
        }
    }

    public bool HasOptionFolder => OptionFolderPath.Length > 0;

    /// <summary>
    /// Варианты запуска задачи - например, пересчитать всё или только часть. Пусто или
    /// один вариант - выбора нет, и в окне он не показывается. Выбор не запоминается между
    /// запусками программы: неполный пересчёт по ошибке хуже лишнего щелчка.
    /// </summary>
    public virtual IReadOnlyList<string> RunOptions => Array.Empty<string>();

    public virtual string RunOptionsTitle => string.Empty;

    public bool HasRunOptions => RunOptions.Count > 1;

    public string? SelectedRunOption
    {
        get => _selectedRunOption ?? RunOptions.FirstOrDefault();
        set
        {
            if (IsBusy || value is null)
            {
                // Во время обработки выбор не меняется: обработчик уже прочитал его.
                OnPropertyChanged();
                return;
            }

            SetProperty(ref _selectedRunOption, value);
            OnPropertyChanged(nameof(RunOptionHint));
            OnPropertyChanged(nameof(ActionCaption));
        }
    }

    /// <summary>Что сделает выбранный вариант.</summary>
    public virtual string RunOptionHint => string.Empty;

    public RelayCommand SelectOptionFolderCommand { get; }

    public RelayCommand ClearOptionFolderCommand { get; }

    public RelayCommand SelectOptionFileCommand { get; }

    public RelayCommand SelectFileCommand { get; }

    public RelayCommand ProcessCommand { get; }

    public RelayCommand CancelCommand { get; }

    public RelayCommand OpenResultFolderCommand { get; }

    /// <summary>Открыть готовую книгу в Excel - обычно это и есть следующий шаг работы.</summary>
    public RelayCommand OpenInExcelCommand { get; }

    /// <summary>Открыть готовую книгу на том месте, о котором говорит замечание.</summary>
    public RelayCommand<ProcessingWarning> OpenWarningCommand { get; }

    public ObservableCollection<ProcessingCounter> Counters { get; } = new();

    public ObservableCollection<ProcessingWarning> Warnings { get; } = new();

    public bool HasWarnings => Warnings.Count > 0;

    /// <summary>Что осмотр нашёл в выбранной книге. Пусто, пока файл не выбран.</summary>
    public ObservableCollection<PreflightIssue> Preflight { get; } = new();

    public bool HasPreflight => Preflight.Count > 0;

    /// <summary>Осмотр нашёл помеху: обработка почти наверняка не дойдёт до конца.</summary>
    public bool HasPreflightProblem => Preflight.Any(issue => issue.IsProblem);

    /// <summary>Книга осмотрена и всё на месте - короткая строчка вместо пустоты.</summary>
    public bool IsPreflightClean => HasSelectedFile && !HasPreflight && _preflightDone;

    /// <summary>
    /// Есть замечания, по которым можно перейти в книгу. Нужно ради подсказки под
    /// списком: обещать переход, когда ни одно замечание к месту не привязано, - обман.
    /// </summary>
    public bool HasNavigableWarnings => Warnings.Any(warning => warning.CanNavigate);

    public string SelectedFilePath
    {
        get => _selectedFilePath;
        private set
        {
            SetProperty(ref _selectedFilePath, value);
            OnPropertyChanged(nameof(HasSelectedFile));
            RefreshCommands();
        }
    }

    public bool HasSelectedFile => SelectedFilePath.Length > 0;

    public string StatusMessage
    {
        get => _statusMessage;
        protected set => SetProperty(ref _statusMessage, value);
    }

    public string ErrorMessage
    {
        get => _errorMessage;
        protected set
        {
            SetProperty(ref _errorMessage, value);
            OnPropertyChanged(nameof(HasError));
            OnPropertyChanged(nameof(ShowHint));
        }
    }

    public bool HasError => ErrorMessage.Length > 0;

    public string ResultPath
    {
        get => _resultPath;
        private set
        {
            SetProperty(ref _resultPath, value);
            OnPropertyChanged(nameof(HasResult));
            OnPropertyChanged(nameof(ShowHint));
            OpenResultFolderCommand.RaiseCanExecuteChanged();
            OpenInExcelCommand.RaiseCanExecuteChanged();
            OpenWarningCommand.RaiseCanExecuteChanged();
        }
    }

    public bool HasResult => ResultPath.Length > 0;

    /// <summary>Подсказка на месте будущего результата: пока ничего не показано.</summary>
    public bool ShowHint => !HasResult && !HasError;

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            SetProperty(ref _isBusy, value);
            OnPropertyChanged(nameof(IsIdle));
            RefreshCommands();
        }
    }

    public bool IsIdle => !IsBusy;

    /// <summary>Над окном держат файл: поле выбора подсвечивается.</summary>
    public bool IsDragOver
    {
        get => _isDragOver;
        set => SetProperty(ref _isDragOver, value);
    }

    public int ProgressValue
    {
        get => _progressValue;
        private set => SetProperty(ref _progressValue, value);
    }

    /// <summary>
    /// Короткая причина сбоя для окна: тип и первая строка сообщения самой глубокой
    /// ошибки - обёртки вроде «исключение в вызванном методе» ничего не объясняют.
    /// </summary>
    private static string ShortReason(Exception ex)
    {
        var inner = ex;
        while (inner.InnerException is not null)
        {
            inner = inner.InnerException;
        }

        var message = (inner.Message ?? string.Empty).Split('\n')[0].Trim();
        if (message.Length > 300)
        {
            message = message[..300] + "…";
        }

        return inner.GetType().Name + (message.Length > 0 ? ": " + message : string.Empty);
    }

    public void Cancel()
    {
        if (_cancellation is null || _cancellation.IsCancellationRequested)
        {
            return;
        }

        StatusMessage = "Отмена обработки...";
        _cancellation.Cancel();
        RefreshCommands();
    }

    /// <summary>
    /// Общий вход для диалога выбора и для файла, перетащенного в окно.
    /// Возвращает false, если формат не поддерживается.
    /// </summary>
    public bool ApplySelectedFile(string path)
    {
        if (IsBusy)
        {
            return false;
        }

        if (!WorkbookFile.IsSupported(path))
        {
            ErrorMessage =
                "Это не книга Excel: " + Path.GetFileName(path) + "." + Environment.NewLine +
                "Подойдут файлы " + string.Join(", ", WorkbookFile.SupportedExtensions) + ".";
            StatusMessage = InitialHint;
            return false;
        }

        if (!File.Exists(path))
        {
            ErrorMessage = "Файл не найден: " + path;
            StatusMessage = InitialHint;
            return false;
        }

        SelectedFilePath = path;
        ErrorMessage = string.Empty;
        ResultPath = string.Empty;
        ProgressValue = 0;
        Counters.Clear();
        ClearWarnings();
        StatusMessage = "Файл выбран. Нажмите «" + ActionCaption + "».";

        RememberFolder(path);
        _ = InspectAsync(path);
        return true;
    }

    private void SelectFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выбор файла",
            Filter = WorkbookFile.DialogFilter,
            CheckFileExists = true,
        };

        // Диалог открывается там же, где книгу брали в прошлый раз: у каждой задачи
        // своя папка, и ходить до неё заново каждый день незачем.
        var folder = _settings?.Get(FolderKey + SettingsKey);
        if (folder is not null && Directory.Exists(folder))
        {
            dialog.InitialDirectory = folder;
        }

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        ApplySelectedFile(dialog.FileName);
    }

    private void SelectOptionFolder()
    {
        var dialog = new OpenFolderDialog { Title = FolderOptionTitle };
        var initial = OptionInitialDirectory();
        if (initial is not null)
        {
            dialog.InitialDirectory = initial;
        }

        if (dialog.ShowDialog() == true)
        {
            OptionFolderPath = dialog.FolderName;
        }
    }

    private void SelectOptionFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = FolderOptionTitle,
            Filter = WorkbookFile.DialogFilter,
            CheckFileExists = true,
        };

        var initial = OptionInitialDirectory();
        if (initial is not null)
        {
            dialog.InitialDirectory = initial;
        }

        if (dialog.ShowDialog() == true)
        {
            OptionFolderPath = dialog.FileName;
        }
    }

    /// <summary>Диалог открывается там, где выбирали в прошлый раз: в папке или в папке выбранного файла.</summary>
    private string? OptionInitialDirectory()
    {
        if (Directory.Exists(OptionFolderPath))
        {
            return OptionFolderPath;
        }

        var parent = OptionFolderPath.Length > 0 ? Path.GetDirectoryName(OptionFolderPath) : null;
        return parent is not null && Directory.Exists(parent) ? parent : null;
    }

    private void RememberFolder(string path)
    {
        var folder = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(folder))
        {
            _settings?.Set(FolderKey + SettingsKey, folder);
        }
    }

    /// <summary>
    /// Осмотр книги: на месте ли листы и доступны ли связанные файлы. Читается zip
    /// книги, а не Excel, поэтому это доли секунды даже на книге в 76 МБ - и человек
    /// узнаёт о непорядке до того, как запустит обработку на полминуты.
    /// </summary>
    private async Task InspectAsync(string path)
    {
        ClearPreflight();

        var issues = await Task.Run(() =>
        {
            var probe = WorkbookProbeReader.Read(path);
            return probe is null
                ? Array.Empty<PreflightIssue>()
                : PreflightCheck.Run(probe, RequiredSheets, File.Exists, OptionalSheets, RepeatableSheets).ToArray();
        }).ConfigureAwait(true);

        // За время осмотра могли выбрать другой файл - тогда этот ответ уже не нужен.
        if (!string.Equals(SelectedFilePath, path, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        foreach (var issue in issues)
        {
            Preflight.Add(issue);
        }

        _preflightDone = true;
        RaisePreflightChanged();

        if (issues.Length > 0)
        {
            _logger.Information(
                "Осмотр книги " + Path.GetFileName(path) + ": " + issues.Length + " замечание(й).");
        }
    }

    private void ClearPreflight()
    {
        _preflightDone = false;
        Preflight.Clear();
        RaisePreflightChanged();
    }

    private void RaisePreflightChanged()
    {
        OnPropertyChanged(nameof(HasPreflight));
        OnPropertyChanged(nameof(HasPreflightProblem));
        OnPropertyChanged(nameof(IsPreflightClean));
    }

    private async Task RunAsync()
    {
        if (IsBusy || SelectedFilePath.Length == 0)
        {
            return;
        }

        IsBusy = true;
        ErrorMessage = string.Empty;
        ResultPath = string.Empty;
        ProgressValue = 0;
        Counters.Clear();
        ClearWarnings();
        StatusMessage = "Подготовка...";

        _cancellation?.Dispose();
        _cancellation = new CancellationTokenSource();

        var progress = new Progress<ProcessingStage>(stage =>
        {
            StatusMessage = stage.Message;
            ProgressValue = stage.Percent;
        });

        try
        {
            var result = await _processor
                .ProcessAsync(SelectedFilePath, progress, _cancellation.Token)
                .ConfigureAwait(true);

            foreach (var counter in result.Counters)
            {
                Counters.Add(counter);
            }

            foreach (var warning in result.Warnings)
            {
                Warnings.Add(warning);
            }

            OnPropertyChanged(nameof(HasWarnings));
            OnPropertyChanged(nameof(HasNavigableWarnings));

            ResultPath = result.OutputPath;
            StatusMessage = SuccessMessage;
            ProgressValue = 100;
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Обработка отменена.";
            ProgressValue = 0;
        }
        catch (WarehousePlanException ex)
        {
            _logger.Error("Ошибка обработки файла.", ex);
            ErrorMessage = ex.Message;
            StatusMessage = "Обработка не выполнена.";
            ProgressValue = 0;
        }
        catch (Exception ex)
        {
            _logger.Error("Непредвиденная ошибка обработки файла.", ex);

            // Причина видна прямо в окне: журнал лежит на рабочем компьютере и к вечеру
            // очищается, а строку из окна легко переслать снимком экрана.
            ErrorMessage =
                FailureMessage + Environment.NewLine +
                "Причина: " + ShortReason(ex) + Environment.NewLine +
                "Подробности записаны в журнал: " + FileAppLogger.DefaultDirectory();
            StatusMessage = "Обработка не выполнена.";
            ProgressValue = 0;
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void ClearWarnings()
    {
        Warnings.Clear();
        OnPropertyChanged(nameof(HasWarnings));
        OnPropertyChanged(nameof(HasNavigableWarnings));
    }

    /// <summary>
    /// Открывает готовую книгу и ставит курсор на ячейку замечания. Работа идёт через
    /// Excel и занимает секунды, поэтому она асинхронная: окно на это время не замирает.
    /// </summary>
    /// <summary>Открывает готовую книгу целиком, без перехода на конкретную ячейку.</summary>
    private Task OpenInExcelAsync() => OpenAsync(null, null);

    private Task OpenWarningAsync(ProcessingWarning warning) =>
        warning.CanNavigate ? OpenAsync(warning.Sheet, warning.Cell) : Task.CompletedTask;

    /// <summary>
    /// Открывает готовую книгу: целиком или сразу на нужной ячейке.
    ///
    /// Защёлка нужна из-за двойного щелчка: по строке замечания это два нажатия кнопки,
    /// и второе успело бы запустить свой Excel, пока первый ещё не зарегистрировался
    /// и не находится.
    /// </summary>
    private async Task OpenAsync(string? sheet, string? cell)
    {
        if (_isNavigating || ResultPath.Length == 0)
        {
            return;
        }

        _isNavigating = true;
        OpenInExcelCommand.RaiseCanExecuteChanged();

        try
        {
            await WorkbookNavigator
                .ShowAsync(ResultPath, sheet, cell, _logger)
                .ConfigureAwait(true);
        }
        catch (WarehousePlanException ex)
        {
            _logger.Warning("Не удалось открыть книгу.", ex);
            ErrorMessage = ex.Message;
        }
        catch (Exception ex)
        {
            _logger.Error("Непредвиденная ошибка при открытии книги.", ex);
            ErrorMessage = "Не удалось открыть книгу. Откройте файл вручную.";
        }
        finally
        {
            _isNavigating = false;
            OpenInExcelCommand.RaiseCanExecuteChanged();
        }
    }

    private void OpenResultFolder()
    {
        if (ResultPath.Length == 0)
        {
            return;
        }

        try
        {
            if (File.Exists(ResultPath))
            {
                Process.Start(new ProcessStartInfo("explorer.exe", "/select,\"" + ResultPath + "\"")
                {
                    UseShellExecute = true,
                });
                return;
            }

            var directory = Path.GetDirectoryName(ResultPath);
            if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
            {
                Process.Start(new ProcessStartInfo(directory) { UseShellExecute = true });
            }
        }
        catch (System.ComponentModel.Win32Exception ex)
        {
            _logger.Warning("Не удалось открыть папку с результатом.", ex);
            ErrorMessage = "Не удалось открыть папку с результатом.";
        }
    }

    private void RefreshCommands()
    {
        SelectFileCommand.RaiseCanExecuteChanged();
        SelectOptionFolderCommand.RaiseCanExecuteChanged();
        SelectOptionFileCommand.RaiseCanExecuteChanged();
        ClearOptionFolderCommand.RaiseCanExecuteChanged();
        ProcessCommand.RaiseCanExecuteChanged();
        CancelCommand.RaiseCanExecuteChanged();
        OpenResultFolderCommand.RaiseCanExecuteChanged();
        OpenInExcelCommand.RaiseCanExecuteChanged();
        OpenWarningCommand.RaiseCanExecuteChanged();
    }
}
