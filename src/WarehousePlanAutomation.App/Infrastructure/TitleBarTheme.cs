using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace WarehousePlanAutomation.App.Infrastructure;

/// <summary>
/// Заголовок окна в цвет темы.
///
/// Заголовок рисует не приложение, а сама Windows, и по умолчанию он остаётся светлым
/// при любой теме окна. Перекрасить его можно только попросив об этом систему -
/// через DwmSetWindowAttribute.
///
/// Начиная с Windows 11 (сборка 22000) цвет заголовка задаётся точно, и он совпадает
/// с полем окна. На Windows 10 такого нет, там доступен только переключатель «тёмный
/// заголовок» - это грубее, но лучше, чем белая полоса над тёмным окном.
/// </summary>
public static class TitleBarTheme
{
    private const int UseImmersiveDarkMode = 20;
    private const int CaptionColor = 35;
    private const int TextColor = 36;
    private const int BorderColor = 34;

    /// <summary>Сборка Windows 11, с которой заголовок можно красить в произвольный цвет.</summary>
    private const int Windows11Build = 22000;

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        IntPtr window, int attribute, ref int value, int size);

    /// <summary>
    /// Красит заголовок окна цветами текущей темы. Так делает и главное окно,
    /// и модальные - иначе над тёмным окном висела бы светлая системная полоса.
    /// </summary>
    public static void ApplyCurrent(Window window)
    {
        Color Of(string key, Color fallback) =>
            window.TryFindResource(key) is SolidColorBrush brush ? brush.Color : fallback;

        // Тёмная тема узнаётся по самому фону окна, а не по настройке: тогда заголовок
        // не может разойтись с палитрой, которая на окне сейчас на самом деле.
        var ground = Of("GroundBrush", Colors.White);

        Apply(
            window,
            IsDark(ground),
            ground,
            Of("InkBrush", IsDark(ground) ? Colors.White : Colors.Black),
            Of("CardBorderBrush", ground));
    }

    /// <summary>Светлота по восприятию: зелёный кажется светлее синего при том же числе.</summary>
    private static bool IsDark(Color color) =>
        ((0.299 * color.R) + (0.587 * color.G) + (0.114 * color.B)) < 128;

    /// <summary>
    /// Красит заголовок окна. Вызывается после появления у окна дескриптора
    /// и при каждой смене темы.
    /// </summary>
    public static void Apply(Window window, bool dark, Color caption, Color text, Color border)
    {
        var handle = new WindowInteropHelper(window).Handle;
        if (handle == IntPtr.Zero)
        {
            return;
        }

        // Тёмный заголовок ставится в любом случае: на Windows 10 это единственное,
        // что доступно, а на Windows 11 он ещё и задаёт цвет кнопок «свернуть/закрыть».
        var immersive = dark ? 1 : 0;
        Set(handle, UseImmersiveDarkMode, ref immersive);

        if (Environment.OSVersion.Version.Build < Windows11Build)
        {
            return;
        }

        var captionValue = ToColorRef(caption);
        var textValue = ToColorRef(text);
        var borderValue = ToColorRef(border);

        Set(handle, CaptionColor, ref captionValue);
        Set(handle, TextColor, ref textValue);
        Set(handle, BorderColor, ref borderValue);
    }

    /// <summary>
    /// Ошибка здесь не должна ничего ломать: не получилось перекрасить заголовок -
    /// окно просто останется с системным. Ради этого не стоит валить приложение.
    /// </summary>
    private static void Set(IntPtr handle, int attribute, ref int value)
    {
        try
        {
            DwmSetWindowAttribute(handle, attribute, ref value, sizeof(int));
        }
        catch (DllNotFoundException)
        {
        }
        catch (EntryPointNotFoundException)
        {
        }
    }

    /// <summary>Windows ждёт цвет в порядке 0x00BBGGRR, а не привычного RGB.</summary>
    private static int ToColorRef(Color color) =>
        color.R | (color.G << 8) | (color.B << 16);
}
