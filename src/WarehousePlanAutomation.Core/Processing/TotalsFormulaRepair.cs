using System.Globalization;
using System.Text.RegularExpressions;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Растягивает формулы под таблицей на фактические строки данных.
///
/// Excel растягивает диапазон сам, только когда строки вставлены внутрь него. Строки
/// добавляются под последней, то есть за границей диапазона, и итог остался бы считаться
/// по старым строкам: после добавления двух строк «=SUBTOTAL(9;P2:P2)» так и продолжил бы
/// показывать одну.
/// </summary>
public static class TotalsFormulaRepair
{
    /// <summary>Диапазон на этом же листе: перед ним нет имени листа с «!».</summary>
    private static readonly Regex LocalRangePattern = new(
        @"(?<![!\w$.])(?<c1>\$?[A-Za-z]{1,3})(?<r1>\$?\d+)\s*:\s*(?<c2>\$?[A-Za-z]{1,3})(?<r2>\$?\d+)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Формула под таблицей: диапазоны, которые начинаются в первой строке данных (или выше,
    /// в заголовке) и кончаются внутри таблицы, растягиваются до последней строки данных.
    /// Так их растянул бы человек после добавления строк. Диапазоны, которые захватывают
    /// строки под таблицей или смотрят на другой лист, не трогаются.
    /// Возвращает null, если менять нечего.
    /// </summary>
    public static string? ExtendBelow(string? formula, int firstDataRow, int lastDataRow, int totalsRow)
    {
        if (string.IsNullOrWhiteSpace(formula) ||
            !formula.StartsWith("=", StringComparison.Ordinal) ||
            lastDataRow < firstDataRow)
        {
            return null;
        }

        var changed = false;
        var result = LocalRangePattern.Replace(formula, match =>
        {
            var row1 = int.Parse(match.Groups["r1"].Value.TrimStart('$'), CultureInfo.InvariantCulture);
            var row2 = int.Parse(match.Groups["r2"].Value.TrimStart('$'), CultureInfo.InvariantCulture);

            if (row1 < 1 || row1 > firstDataRow || row2 < firstDataRow || row2 >= totalsRow || row2 == lastDataRow)
            {
                return match.Value;
            }

            changed = true;
            return match.Groups["c1"].Value + match.Groups["r1"].Value + ":" +
                   match.Groups["c2"].Value + Row(match.Groups["r2"].Value, lastDataRow);
        });

        return changed ? result : null;
    }

    /// <summary>Знак доллара сохраняется таким, каким его написал человек.</summary>
    private static string Row(string original, int row) =>
        (original.StartsWith("$", StringComparison.Ordinal) ? "$" : string.Empty) +
        row.ToString(CultureInfo.InvariantCulture);
}
