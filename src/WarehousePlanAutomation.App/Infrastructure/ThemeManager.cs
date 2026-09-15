using System.IO;
using System.Windows;
using WarehousePlanAutomation.Core.Logging;

namespace WarehousePlanAutomation.App.Infrastructure;

public enum AppTheme
{
    Light,
    Dark,
}

/// <summary>
/// Подмена палитры на ходу.
///
/// Цвета лежат отдельными словарями (<c>Themes/Light.xaml</c> и <c>Themes/Dark.xaml</c>)
/// с одинаковыми ключами, поэтому переключение - это замена одного словаря на другой.
/// Работает это только потому, что стили ссылаются на кисти через DynamicResource:
/// StaticResource запомнил бы цвет один раз, при разборе разметки.
///
/// Выбор запоминается рядом с журналом работы: программа переносится одним файлом,
/// и место для настроек у неё то же самое.
/// </summary>
public sealed class ThemeManager
{
    private const string LightSource = "Themes/Light.xaml";
    private const string DarkSource = "Themes/Dark.xaml";

    /// <summary>
    /// Адрес словаря пишется полностью, с именем сборки. Относительный путь
    /// разрешается от запускающей сборки, а не от той, где словарь лежит, - и стоит
    /// собрать окно из другого приложения (например, из харнесса снимков), как
    /// относительный адрес перестаёт находиться.
    /// </summary>
    private static Uri SourceOf(string relative) =>
        new("pack://application:,,,/" + typeof(ThemeManager).Assembly.GetName().Name +
            ";component/" + relative,
            UriKind.Absolute);

    private readonly IAppLogger _logger;
    private readonly string _settingsPath;

    public ThemeManager(IAppLogger logger)
        : this(logger, DefaultSettingsPath())
    {
    }

    public ThemeManager(IAppLogger logger, string settingsPath)
    {
        _logger = logger;
        _settingsPath = settingsPath;
        Current = Read();
    }

    public AppTheme Current { get; private set; }

    public bool IsDark => Current == AppTheme.Dark;

    /// <summary>Ставит запомненную тему. Вызывается один раз при запуске.</summary>
    public void ApplySaved() => Apply(Current, save: false);

    public void Toggle() => Apply(IsDark ? AppTheme.Light : AppTheme.Dark, save: true);

    private void Apply(AppTheme theme, bool save)
    {
        var resources = Application.Current?.Resources;
        if (resources is null)
        {
            return;
        }

        var wantedSource = theme == AppTheme.Dark ? DarkSource : LightSource;
        var merged = resources.MergedDictionaries;
        var index = IndexOfPalette(merged);

        // Уже стоит нужный словарь - трогать ничего не нужно. Это не только экономия:
        // при запуске со светлой темой словари приложения вообще не меняются.
        if (index >= 0 && IsSource(merged[index], wantedSource))
        {
            Current = theme;
            return;
        }

        var wanted = new ResourceDictionary { Source = SourceOf(wantedSource) };

        // Словарь заменяется на месте, а не удаляется и вставляется заново: удаление
        // с последующей вставкой не доходит до окон, созданных сразу после него, и окно
        // рисуется прежней палитрой. Замена по индексу оповещает об изменении как надо.
        if (index >= 0)
        {
            merged[index] = wanted;
        }
        else
        {
            merged.Insert(0, wanted);
        }

        Current = theme;

        if (save)
        {
            Write(theme);
        }

        _logger.Information("Тема оформления: " + (theme == AppTheme.Dark ? "тёмная" : "светлая") + ".");
    }

    /// <summary>Где в списке словарей лежит палитра. -1, если её там ещё нет.</summary>
    private static int IndexOfPalette(IList<ResourceDictionary> merged)
    {
        for (var i = 0; i < merged.Count; i++)
        {
            if (IsSource(merged[i], LightSource) || IsSource(merged[i], DarkSource))
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Словарь приложения задан относительным адресом, подставленный - полным,
    /// поэтому сравнивается окончание.
    /// </summary>
    private static bool IsSource(ResourceDictionary dictionary, string relative) =>
        dictionary.Source?.OriginalString.EndsWith(relative, StringComparison.OrdinalIgnoreCase) == true;

    private AppTheme Read()
    {
        try
        {
            return File.Exists(_settingsPath) &&
                   File.ReadAllText(_settingsPath).Trim()
                       .Equals("dark", StringComparison.OrdinalIgnoreCase)
                ? AppTheme.Dark
                : AppTheme.Light;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _logger.Warning("Не удалось прочитать выбранную тему, взята светлая.", ex);
            return AppTheme.Light;
        }
    }

    private void Write(AppTheme theme)
    {
        try
        {
            var directory = Path.GetDirectoryName(_settingsPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(_settingsPath, theme == AppTheme.Dark ? "dark" : "light");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Не запомнить тему - мелочь: в этот раз она уже применена.
            _logger.Warning("Не удалось запомнить выбранную тему.", ex);
        }
    }

    private static string DefaultSettingsPath() =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WarehousePlanAutomation",
            "theme.txt");
}
