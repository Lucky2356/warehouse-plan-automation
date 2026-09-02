using System.Globalization;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Разбор листа «Журнал заказов на отгрузку» для одного номера загрузки.
///
/// «% выполнения» считается как отношение сумм: сколько единиц собрано фактически
/// к тому, сколько заказано. Колонка «%» самого журнала для этого не годится -
/// в ней процент отдельного документа, а не заказа целиком.
///
/// Два обстоятельства делают простое суммирование неверным:
///
/// 1. Часть строк журнала - сводные: в колонке «Номер» у них стоит не документ
///    («З000-260355»), а число («-21»). Их «Кол-во ед» уже включает строки
///    магазинов, поэтому при суммировании всё удвоилось бы.
/// 2. Лист журнала содержит несколько подряд приклеенных выгрузок, и один и тот же
///    документ встречается в нём по нескольку раз. В разобранных файлах повторы
///    полностью совпадают по числам, но старые заказы повторяются чаще новых,
///    поэтому без дедупликации итог перекашивается в их пользу.
///
/// Отсюда правило: считаем только строки-документы, каждый номер документа - один раз.
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
            .ToList();

        if (occurrences.Count == 0)
        {
            return JournalOutcome.NotFound;
        }

        var percent = CalculatePercent(occurrences);

        // Статус «Запущен» означает, что заказ уже в сборке, даже если собрано ноль единиц.
        var hasStarted = occurrences
            .Any(row => TextUtils.EqualsKey(row.Status, OrderTextRules.JournalStartedStatus));

        return new JournalOutcome(true, percent > 0d || hasStarted, percent);
    }

    private static double CalculatePercent(IReadOnlyList<JournalRow> occurrences)
    {
        var counted = new HashSet<string>(StringComparer.Ordinal);
        var planned = 0d;
        var actual = 0d;

        foreach (var row in occurrences)
        {
            if (!IsDocumentRow(row.DocumentNumber))
            {
                continue;
            }

            if (!counted.Add(TextUtils.NormalizeKey(row.DocumentNumber)))
            {
                continue;
            }

            planned += row.Quantity ?? 0d;
            actual += row.ActualQuantity ?? 0d;
        }

        if (planned <= 0d)
        {
            return 0d;
        }

        return Math.Round(actual / planned * 100d, MidpointRounding.AwayFromZero);
    }

    /// <summary>
    /// Строка документа магазина, а не сводная. Отличаются они тем, что у сводных
    /// в колонке «Номер» стоит число.
    /// </summary>
    private static bool IsDocumentRow(string? documentNumber)
    {
        var text = TextUtils.Normalize(documentNumber);
        return text.Length > 0 && TextUtils.CellToDouble(text) is null;
    }
}
