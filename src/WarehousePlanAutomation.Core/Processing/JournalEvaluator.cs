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
        // Номер ищется с границами по цифрам: «55575395» не должен находиться внутри
        // более длинного числа вроде «155575395», иначе заказу достался бы чужой статус.
        var key = loadNumber.ToString(CultureInfo.InvariantCulture);
        var occurrences = journal
            .Where(row => TextUtils.ContainsNumber(row.Comment, key))
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

        // Первое вхождение тоже проверяется на «Запущен»: строка со статусом «Запущен»
        // и нулевым процентом означает, что заказ уже в сборке, независимо от того,
        // первая она в журнале или нет.
        var hasStarted = occurrences
            .Any(row => TextUtils.EqualsKey(row.Status, OrderTextRules.JournalStartedStatus));

        return new JournalOutcome(true, hasStarted, null);
    }
}
