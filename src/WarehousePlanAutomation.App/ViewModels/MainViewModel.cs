using System.Diagnostics;
using System.Globalization;
using System.IO;
using Microsoft.Win32;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;

namespace WarehousePlanAutomation.App.ViewModels;

public sealed class MainViewModel : ObservableObject
{
    private readonly IWorkbookProcessor _processor;
    private readonly IAppLogger _logger;

    private string _selectedFilePath = string.Empty;
    private string _statusMessage = "Выберите Excel-файл с листами «Заказы на отгрузку» и «Журнал заказов на отгрузку».";
    private string _errorMessage = string.Empty;
    private string _resultPath = string.Empty;
    private string _resultSummary = string.Empty;
    private bool _isBusy;
    private int _progressValue;

    private CancellationTokenSource? _cancellation;

    public MainViewModel(IWorkbookProcessor processor, IAppLogger logger)
    {
        _processor = processor;
        _logger = logger;

        SelectFileCommand = new RelayCommand(SelectFile, () => !IsBusy);
        ProcessCommand = new RelayCommand(
            () => _ = RunAsync(),
            () => !IsBusy && SelectedFilePath.Length > 0);
        OpenResultFolderCommand = new RelayCommand(OpenResultFolder, () => ResultPath.Length > 0);
    }

    public RelayCommand SelectFileCommand { get; }

    public RelayCommand ProcessCommand { get; }

    public RelayCommand OpenResultFolderCommand { get; }

    public string SelectedFilePath
    {
        get => _selectedFilePath;
        private set
        {
            SetProperty(ref _selectedFilePath, value);
            RefreshCommands();
        }
    }

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
            OpenResultFolderCommand.RaiseCanExecuteChanged();
        }
    }

    public bool HasResult => ResultPath.Length > 0;

    public string ResultSummary
    {
        get => _resultSummary;
        private set => SetProperty(ref _resultSummary, value);
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

    public int ProgressValue
    {
        get => _progressValue;
        private set => SetProperty(ref _progressValue, value);
    }

    public void Cancel() => _cancellation?.Cancel();

    private void SelectFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выбор файла",
            Filter = "Книги Excel (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|Все файлы (*.*)|*.*",
            CheckFileExists = true,
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        SelectedFilePath = dialog.FileName;
        ErrorMessage = string.Empty;
        ResultPath = string.Empty;
        ResultSummary = string.Empty;
        ProgressValue = 0;
        StatusMessage = "Файл выбран. Нажмите «Сформировать план склада».";
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
        ResultSummary = string.Empty;
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

            ResultPath = result.OutputPath;
            ResultSummary = BuildSummary(result);
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

    private static string BuildSummary(ProcessingResult result)
    {
        var culture = CultureInfo.CurrentCulture;
        return
            "Обработано строк выгрузки: " + result.ProcessedOrderRows.ToString(culture) + Environment.NewLine +
            "Осталось строк для ручной проверки: " + result.RemainingOrderRows.ToString(culture) + Environment.NewLine +
            "Добавлено новых заказов: " + result.NewPlanOrders.ToString(culture) + Environment.NewLine +
            "Удалено строк «Заказы будут загружены»: " + result.DeletedPlaceholderRows.ToString(culture);
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
        OpenResultFolderCommand.RaiseCanExecuteChanged();
    }
}
