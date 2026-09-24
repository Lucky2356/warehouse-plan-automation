using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Лист «Исключения»: секторы, группы и АЦР, которые отдаём без согласования, что бы
/// ни стояло в «Запрете». Листа в книге нет - исключений нет вовсе.
///
/// Аналитик ведёт лист свободно, поэтому сравнивается всё, что на нём написано: значение
/// подходит, если совпало с сектором, группой, артикулом или АЦР строки. АЦР сверяется ещё
/// и началом - так «1235-190ЗОЛОТО» закрывает все размеры этого цвета.
/// </summary>
public sealed class RestockExceptions
{
    private readonly HashSet<string> _values;

    public RestockExceptions(IEnumerable<string> values)
    {
        _values = values
            .Select(TextUtils.NormalizeKey)
            .Where(value => value.Length > 0)
            .ToHashSet(StringComparer.Ordinal);
    }

    public static RestockExceptions Empty { get; } = new(Array.Empty<string>());

    public int Count => _values.Count;

    public bool Covers(string? sector, string? group, string? article, string? acr)
    {
        if (_values.Count == 0)
        {
            return false;
        }

        foreach (var value in new[] { sector, group, article, acr })
        {
            var key = TextUtils.NormalizeKey(value);
            if (key.Length > 0 && _values.Contains(key))
            {
                return true;
            }
        }

        // Начало АЦР сверяется только с длинными значениями: короткое «5» подошло бы
        // к половине книги.
        var acrKey = TextUtils.NormalizeKey(acr);
        return acrKey.Length > 0 &&
               _values.Any(value => value.Length >= 5 && acrKey.StartsWith(value, StringComparison.Ordinal));
    }
}
