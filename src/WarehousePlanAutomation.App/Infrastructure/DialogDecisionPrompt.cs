using System.Windows;
using WarehousePlanAutomation.App.Views;
using WarehousePlanAutomation.Core.Abstractions;

namespace WarehousePlanAutomation.App.Infrastructure;

/// <summary>
/// Задаёт вопрос человеку модальным окном.
///
/// Обработка идёт на своём STA-потоке, поэтому вызов переключается на поток интерфейса
/// через <see cref="System.Windows.Threading.Dispatcher.Invoke(Action)"/> и там же блокируется
/// до ответа. Для потока обработки это безопасно: все вызовы COM исходящие,
/// входящих обратных вызовов из Excel нет.
/// </summary>
public sealed class DialogDecisionPrompt : IDecisionPrompt
{
    private readonly Window _owner;

    public DialogDecisionPrompt(Window owner)
    {
        _owner = owner;
    }

    public string? Choose(DecisionRequest request)
    {
        if (request.Options.Count == 0)
        {
            return null;
        }

        return _owner.Dispatcher.Invoke(() =>
        {
            var window = new DecisionWindow(request) { Owner = _owner };
            return window.ShowDialog() == true ? window.Choice : null;
        });
    }
}
