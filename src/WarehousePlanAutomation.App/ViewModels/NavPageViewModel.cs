using WarehousePlanAutomation.App.Infrastructure;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>
/// Страница боковой панели. Задача обработки книги - её частный случай, но не
/// единственный: обновления книг не открывают, а в списке стоять должны.
/// </summary>
public abstract class NavPageViewModel : ObservableObject
{
    private string _badge = string.Empty;

    /// <summary>Название в боковой панели.</summary>
    public abstract string Title { get; }

    /// <summary>Подпись под заголовком: чем страница занимается.</summary>
    public abstract string Subtitle { get; }

    /// <summary>
    /// Значок страницы - контур в координатах 20×20. Нужен, чтобы на узком окне,
    /// где подписи не помещаются, задачи всё равно различались.
    /// </summary>
    public abstract string IconData { get; }

    /// <summary>
    /// Значок у пункта панели. Нужен, чтобы новое обновление было видно и тогда,
    /// когда открыта другая страница.
    /// </summary>
    public string Badge
    {
        get => _badge;
        protected set
        {
            SetProperty(ref _badge, value);
            OnPropertyChanged(nameof(HasBadge));
        }
    }

    public bool HasBadge => Badge.Length > 0;
}
