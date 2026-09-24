using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WarehousePlanAutomation.App.Views;

/// <summary>Экран одной задачи: файл, запуск, ход обработки и результат.</summary>
public partial class WorkbookTaskView : UserControl
{
    public WorkbookTaskView()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Список вариантов запуска прокручивать некуда - он в одну строку, но колесо мыши
    /// он всё равно перехватывает, и страница под курсором стоит на месте. Событие
    /// отправляется дальше, наверх: крутится вся страница, где бы ни был курсор.
    /// </summary>
    private void OnInnerListWheel(object sender, MouseWheelEventArgs e)
    {
        if (e.Handled || sender is not UIElement source)
        {
            return;
        }

        e.Handled = true;
        source.RaiseEvent(new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
        {
            RoutedEvent = UIElement.MouseWheelEvent,
            Source = source,
        });
    }
}
