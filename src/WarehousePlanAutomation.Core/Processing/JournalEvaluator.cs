using System.Globalization;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Разбор листа «Журнал заказов на отгрузку» для одного номера загрузки.
/// Журнал не сортируется: первым вхождением считается самая верхняя найденная строка.
/// </summary>
public static class JournalEvaluator
{
    public static JournalOutcome Evaluate(long loadNumber, IReadOnlyList<JournalRow> journal)
    {
        var key = loadNumber.ToString(CultureInfo.InvariantCulture);
        var occurrences = journal
            .Where(row => row.Comment.Contains(key, StringComparison.Ordinal))
            .OrderBy(row => row.Order)
            .ToList();

        if (occurrences.Count == 0)
        {
            return JournalOutcome.NotFound;
        }

        var first = occurrences[0];
        var firstPercent = first.Percent ?? 0d;

        if (firstPercent > 0d)
        {
            return new JournalOutcome(true, true, firstPercent);
        }

        var hasStarted = occurrences
            .Skip(1)
            .Any(row => TextUtils.EqualsKey(row.Status, OrderTextRules.JournalStartedStatus));

        return new JournalOutcome(true, hasStarted, null);
    }
}
