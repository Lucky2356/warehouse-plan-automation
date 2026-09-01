using System.Globalization;
using System.Text;

namespace WarehousePlanAutomation.Core.Text;

/// <summary>
/// Нормализация текста ячеек: регистр, начальные/конечные пробелы, повторные пробелы,
/// неразрывные пробелы. Используется всеми правилами поиска заголовков и служебных строк.
/// </summary>
public static class TextUtils
{
    private const char NoBreakSpace = '\u00A0';
    private const char NarrowNoBreakSpace = '\u202F';
    private const char ZeroWidthSpace = '\u200B';

    /// <summary>Схлопывает пробельные символы, убирает неразрывные пробелы и обрезает края.</summary>
    public static string Normalize(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length);
        var pendingSpace = false;
        foreach (var raw in value)
        {
            if (raw == ZeroWidthSpace)
            {
                continue;
            }

            var ch = raw == NoBreakSpace || raw == NarrowNoBreakSpace ? ' ' : raw;
            if (char.IsWhiteSpace(ch))
            {
                if (builder.Length > 0)
                {
                    pendingSpace = true;
                }

                continue;
            }

            if (pendingSpace)
            {
                builder.Append(' ');
                pendingSpace = false;
            }

            builder.Append(ch);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Нормализованный ключ для сравнения без учёта регистра.
    /// Буква «ё» приводится к «е»: в реальных выгрузках одно и то же название
    /// встречается в обоих написаниях (например «Приемка» и «Приёмка»).
    /// </summary>
    public static string NormalizeKey(string? value) =>
        Normalize(value).ToLowerInvariant().Replace('ё', 'е');

    public static bool EqualsKey(string? value, string normalizedNeedle) =>
        string.Equals(NormalizeKey(value), normalizedNeedle, StringComparison.Ordinal);

    public static bool StartsWithKey(string? value, string normalizedNeedle) =>
        NormalizeKey(value).StartsWith(normalizedNeedle, StringComparison.Ordinal);

    public static bool ContainsKey(string? value, string normalizedNeedle) =>
        NormalizeKey(value).Contains(normalizedNeedle, StringComparison.Ordinal);

    /// <summary>
    /// Вхождение как отдельного слова: соседние символы не должны быть буквами или цифрами.
    /// Нужно для коротких названий, которые встречаются внутри других слов:
    /// «Магнит» является маркетплейсом, а «Магнитогорск-М96» - обычный магазин.
    /// </summary>
    public static bool ContainsWord(string? value, string normalizedNeedle)
    {
        if (normalizedNeedle.Length == 0)
        {
            return false;
        }

        var key = NormalizeKey(value);
        var index = key.IndexOf(normalizedNeedle, StringComparison.Ordinal);
        while (index >= 0)
        {
            var startsWord = index == 0 || !IsWordCharacter(key[index - 1]);
            var endIndex = index + normalizedNeedle.Length;
            var endsWord = endIndex >= key.Length || !IsWordCharacter(key[endIndex]);

            if (startsWord && endsWord)
            {
                return true;
            }

            index = key.IndexOf(normalizedNeedle, index + 1, StringComparison.Ordinal);
        }

        return false;
    }

    /// <summary>
    /// Вхождение как отдельного признака: слева граница слова, справа не буква.
    /// Цифры справа допускаются, потому что признак может быть пронумерован.
    /// Так «СЕТ», «СЕТ1» и «СЕТ2» находятся, а «в сети» и «кассета» - нет.
    /// </summary>
    public static bool ContainsToken(string? value, string normalizedNeedle)
    {
        if (normalizedNeedle.Length == 0)
        {
            return false;
        }

        var key = NormalizeKey(value);
        var index = key.IndexOf(normalizedNeedle, StringComparison.Ordinal);
        while (index >= 0)
        {
            var startsWord = index == 0 || !IsWordCharacter(key[index - 1]);
            var endIndex = index + normalizedNeedle.Length;
            var endsToken = endIndex >= key.Length || !char.IsLetter(key[endIndex]);

            if (startsWord && endsToken)
            {
                return true;
            }

            index = key.IndexOf(normalizedNeedle, index + 1, StringComparison.Ordinal);
        }

        return false;
    }

    /// <summary>
    /// Вхождение числа с границами по цифрам: «55575395» не должно находиться
    /// внутри более длинного числа вроде «155575395» или «555753950».
    /// </summary>
    public static bool ContainsNumber(string? value, string digits)
    {
        if (string.IsNullOrEmpty(value) || digits.Length == 0)
        {
            return false;
        }

        var index = value.IndexOf(digits, StringComparison.Ordinal);
        while (index >= 0)
        {
            var startsNumber = index == 0 || !char.IsDigit(value[index - 1]);
            var endIndex = index + digits.Length;
            var endsNumber = endIndex >= value.Length || !char.IsDigit(value[endIndex]);

            if (startsNumber && endsNumber)
            {
                return true;
            }

            index = value.IndexOf(digits, index + 1, StringComparison.Ordinal);
        }

        return false;
    }

    private static bool IsWordCharacter(char value) => char.IsLetterOrDigit(value);

    public static bool ContainsAnyKey(string? value, IEnumerable<string> normalizedNeedles)
    {
        var key = NormalizeKey(value);
        foreach (var needle in normalizedNeedles)
        {
            if (key.Contains(needle, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>Приводит значение ячейки (Value2) к строке без потери длинных чисел.</summary>
    public static string CellToString(object? cell)
    {
        switch (cell)
        {
            case null:
                return string.Empty;
            case string s:
                return s;
            case bool b:
                return b ? "ИСТИНА" : "ЛОЖЬ";
            case double d:
                return d == Math.Floor(d) && Math.Abs(d) < 1e15
                    ? ((long)d).ToString(CultureInfo.InvariantCulture)
                    : d.ToString("R", CultureInfo.InvariantCulture);
            case DateTime dt:
                return dt.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
            case IFormattable formattable:
                return formattable.ToString(null, CultureInfo.InvariantCulture);
            default:
                return cell.ToString() ?? string.Empty;
        }
    }

    /// <summary>Приводит значение ячейки к числу. Строки разбираются и с точкой, и с запятой.</summary>
    public static double? CellToDouble(object? cell)
    {
        switch (cell)
        {
            case null:
                return null;
            case double d:
                return d;
            case int i:
                return i;
            case long l:
                return l;
            case decimal m:
                return (double)m;
            case bool b:
                return b ? 1d : 0d;
            case DateTime dt:
                return dt.ToOADate();
        }

        var text = Normalize(CellToString(cell));
        if (text.Length == 0)
        {
            return null;
        }

        text = text.Replace(" ", string.Empty).Replace(',', '.');
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }
}
