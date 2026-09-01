using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;

namespace WarehousePlanAutomation.App.ViewModels;

public sealed class MainViewModel : ObservableObject
{
    private const string InitialHint =
        "Выберите вчерашнюю книгу с обновлёнными листами «Заказы на отгрузку» " +
        "и «Журнал заказов на отгрузку».";

    private readonly IWorkbookProcessor _processor;
    private readonly IAppLogger _logger;

    private string _selectedFilePath = string.Empty;
    private string _statusMessage = InitialHint;
    private string _errorMessage = string.Empty;
    private string _resultPath = string.Empty;
    private bool _isBusy;
    private bool _isDragOver;
    private int _progressValue;
    private int _processedRows;
    private int _remainingRows;
    private int _newOrders;
    private int _deletedPlaceholders;

    private CancellationTokenSource? _cancellation;

    public MainViewModel(IWorkbookProcessor processor, IAppLogger logger)
    {
        _processor = processor;
        _logger = logger;

        SelectFileCommand = new RelayCommand(SelectFile, () => !IsBusy);
        ProcessCommand = new RelayCommand(
            () => _ = RunAsync(),
            () => !IsBusy && SelectedFilePath.Length > 0);
        CancelCommand = new RelayCommand(Cancel, () => IsBusy);
        OpenResultFolderCommand = new RelayCommand(OpenResultFolder, () => ResultPath.Length > 0);
    }

    public RelayCommand SelectFileCommand { get; }

    public RelayCommand ProcessCommand { get; }

    public RelayCommand CancelCommand { get; }

    public RelayCommand OpenResultFolderCommand { get; }

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
        private set => SetProperty(ref _statusMessage, value);
    }

    public string ErrorMessage
    {
        get => _errorMessage;
        private set
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
        }
    }

    public bool HasResult => ResultPath.Length > 0;

    /// <summary>Подсказка на месте будущего результата: пока ничего не показано.</summary>
    public bool ShowHint => !HasResult && !HasError;

    public int ProcessedRows
    {
        get => _processedRows;
        private set => SetProperty(ref _processedRows, value);
    }

    /// <summary>
    /// Строки, которые не распределились автоматически. Единственный счётчик,
    /// требующий действия: если он больше нуля, окно подсвечивает его отдельно.
    /// </summary>
    public int RemainingRows
    {
        get => _remainingRows;
        private set
        {
            SetProperty(ref _remainingRows, value);
            OnPropertyChanged(nameof(HasLeftovers));
        }
    }

    public bool HasLeftovers => RemainingRows > 0;

    public int NewOrders
    {
        get => _newOrders;
        private set => SetProperty(ref _newOrders, value);
    }

    public int DeletedPlaceholders
    {
        get => _deletedPlaceholders;
        private set => SetProperty(ref _deletedPlaceholders, value);
    }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            SetProperty(ref _isBusy, value);
            RefreshCommands();
        }
    }

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
        StatusMessage = "Файл выбран. Нажмите «Сформировать план склада».";
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

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        ApplySelectedFile(dialog.FileName);
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

            ProcessedRows = result.ProcessedOrderRows;
            RemainingRows = result.RemainingOrderRows;
            NewOrders = result.NewPlanOrders;
            DeletedPlaceholders = result.DeletedPlaceholderRows;
            ResultPath = result.OutputPath;
            StatusMessage = "План склада успешно сформирован";
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
            ErrorMessage =
                "Не удалось сформировать план склада. Подробности записаны в журнал:" + Environment.NewLine +
                FileAppLogger.DefaultDirectory();
            StatusMessage = "Обработка не выполнена.";
            ProgressValue = 0;
        }
        finally
        {
            IsBusy = false;
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
        ProcessCommand.RaiseCanExecuteChanged();
        CancelCommand.RaiseCanExecuteChanged();
        OpenResultFolderCommand.RaiseCanExecuteChanged();
    }
}
