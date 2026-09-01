using System.Windows;
using WarehousePlanAutomation.App.ViewModels;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Excel;

namespace WarehousePlanAutomation.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();

        var logger = ((App)Application.Current).Logger;
        _viewModel = new MainViewModel(new ExcelWorkbookProcessor(logger), logger);
        DataContext = _viewModel;
    }

    protected override void OnClosed(EventArgs e)
    {
        _viewModel.Cancel();
        base.OnClosed(e);
    }

    private void OnDragOver(object sender, DragEventArgs e)
    {
        var accepted = !_viewModel.IsBusy && TryGetWorkbookPath(e, out _);

        e.Effects = accepted ? DragDropEffects.Copy : DragDropEffects.None;
        _viewModel.IsDragOver = accepted;
        e.Handled = true;
    }

    private void OnDragLeave(object sender, DragEventArgs e)
    {
        _viewModel.IsDragOver = false;
        e.Handled = true;
    }

    private void OnDrop(object sender, DragEventArgs e)
    {
        _viewModel.IsDragOver = false;
        e.Handled = true;

        if (TryGetWorkbookPath(e, out var path))
        {
            _viewModel.ApplySelectedFile(path);
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
