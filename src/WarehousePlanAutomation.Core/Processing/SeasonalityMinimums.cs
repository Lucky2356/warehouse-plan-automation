using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// «Мин раздача» блока сезонности листа «Распред». Её вносит человек - по подгруппе.
///
/// Список подгрупп программа пишет заново, и «Мин раздача» прошлого распреда стояла бы
/// напротив чужих подгрупп, а ниже списка - напротив пустых строк. Поэтому значение
/// остаётся только у подгруппы, у которой оно уже было, а всё прошлое убирается.
/// </summary>
public static class SeasonalityMinimums
{
    /// <param name="previous">Подгруппа и «Мин раздача» каждой строки блока до записи.</param>
    /// <param name="subgroups">Новый список подгрупп.</param>
    /// <returns>«Мин раздача» для каждой строки нового списка; null - пусто.</returns>
    public static IReadOnlyList<object?> Carry(
        IReadOnlyList<(string? Subgroup, object? Minimum)> previous,
        IReadOnlyList<string> subgroups)
    {
        var known = new Dictionary<string, object>(StringComparer.Ordinal);
        foreach (var (subgroup, minimum) in previous)
        {
            var key = TextUtils.NormalizeKey(subgroup);
            if (key.Length > 0 && minimum is not null && TextUtils.Normalize(TextUtils.CellToString(minimum)).Length > 0)
            {
                known.TryAdd(key, minimum);
            }
        }

        return subgroups
            .Select(subgroup => known.TryGetValue(TextUtils.NormalizeKey(subgroup), out var value) ? value : null)
            .ToList();
    }
}
