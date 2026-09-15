using System.Collections.ObjectModel;
using System.Reflection;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Updates;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>
/// Оболочка окна: список страниц слева и выбранная страница справа.
/// Кнопки «Назад» нет намеренно - список виден всегда.
/// </summary>
public sealed class ShellViewModel : ObservableObject
{
    private readonly ThemeManager? _theme;

    private NavPageViewModel _selectedPage;
    private bool _isCompact;

    public ShellViewModel(IReadOnlyList<NavPageViewModel> pages, ThemeManager? theme = null)
    {
        _theme = theme;
        ToggleThemeCommand = new RelayCommand(ToggleTheme);

        if (pages.Count == 0)
        {
            throw new ArgumentException("Нужна хотя бы одна страница.", nameof(pages));
        }

        Pages = new ObservableCollection<NavPageViewModel>(pages);
        _selectedPage = Pages[0];
    }

    public ObservableCollection<NavPageViewModel> Pages { get; }

    public NavPageViewModel SelectedPage
    {
        get => _selectedPage;
        set
        {
            if (value is null)
            {
                return;
            }

            SetProperty(ref _selectedPage, value);
            OnPropertyChanged(nameof(SelectedTask));
        }
    }

    /// <summary>
    /// Окно узкое: боковая панель сворачивается в значки. Считает не она сама -
    /// ширину знает только окно, и оно же это свойство и выставляет.
    /// </summary>
    public bool IsCompact
    {
        get => _isCompact;
        set => SetProperty(ref _isCompact, value);
    }

    /// <summary>Версия программы - показывается под её названием в боковой панели.</summary>
    public string Version { get; } = ReadVersion();

    public RelayCommand ToggleThemeCommand { get; }

    /// <summary>Сейчас включена тёмная тема. От неё зависит только значок переключателя.</summary>
    public bool IsDarkTheme => _theme?.IsDark ?? false;

    /// <summary>Подпись переключателя: что произойдёт по нажатию, а не что сейчас.</summary>
    public string ThemeHint => IsDarkTheme ? "Светлая тема" : "Тёмная тема";

    private void ToggleTheme()
    {
        if (_theme is null)
        {
            return;
        }

        _theme.Toggle();
        OnPropertyChanged(nameof(IsDarkTheme));
        OnPropertyChanged(nameof(ThemeHint));
    }

    /// <summary>Задачи обработки книг - к ним относятся выбор файла и перетаскивание.</summary>
    public IEnumerable<WorkbookTaskViewModel> Tasks => Pages.OfType<WorkbookTaskViewModel>();

    /// <summary>Открытая задача обработки или null, если открыта другая страница.</summary>
    public WorkbookTaskViewModel? SelectedTask => SelectedPage as WorkbookTaskViewModel;

    private static string ReadVersion()
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        return version is null ? string.Empty : "версия " + AppVersion.Display(version);
    }
}
