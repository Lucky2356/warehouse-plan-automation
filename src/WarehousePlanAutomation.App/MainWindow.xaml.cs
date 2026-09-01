using System.Windows;
using WarehousePlanAutomation.App.ViewModels;
using WarehousePlanAutomation.Core.Logging;
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
}
