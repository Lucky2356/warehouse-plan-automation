using System.Globalization;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Строка «Склада». <paramref name="Values"/> - значения для колонок «А2, А3» и «МП»
/// в порядке <see cref="Sheets.ReceivingSchema.Storage.FromWarehouse"/>.
/// </summary>
public sealed record WarehouseRow(
    int Index,
    object? StorageType,
    object? Container,
    object? Code,
    object? Quantity,
    IReadOnlyList<object?> Values);

/// <summary>Что осталось после отбора и сколько строк ушло на удалении дубликатов.</summary>
public sealed record StorageSelectionResult(
    IReadOnlyList<WarehouseRow> Rows,
    int Matched,
    int DuplicateContainers,
    int BlankContainers);

/// <summary>
/// Листы «А2, А3» и «МП» со «Склада»: строки одного типа хранения, без повторов тары,
/// по коду и количеству по возрастанию.
///
/// Повторы убираются так же, как это делает «Удалить дубликаты» по колонке «Тара»:
/// остаётся первая строка в порядке «Склада», пустая тара тоже считается значением.
/// Сортировка - как в Excel: числа раньше текста, пустое количество - в конце.
/// Порядок важен: формулы «Учитывать» набирают количество с меньших тар.
/// </summary>
public static class StorageSelection
{
    public static StorageSelectionResult Select(IEnumerable<WarehouseRow> rows, string storageType)
    {
        var wanted = TextUtils.NormalizeKey(storageType);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var kept = new List<WarehouseRow>();
        int matched = 0, duplicates = 0, blanks = 0;

        foreach (var row in rows)
        {
            if (!TextUtils.EqualsKey(TextUtils.CellToString(row.StorageType), wanted))
            {
                continue;
            }

            matched++;
            var container = TextUtils.Normalize(TextUtils.CellToString(row.Container));
            if (!seen.Add(container))
            {
                duplicates++;
                if (container.Length == 0)
                {
                    blanks++;
                }

                continue;
            }

            kept.Add(row);
        }

        var sorted = kept
            .OrderBy(row => row.Code, ExcelOrder.Instance)
            .ThenBy(row => row.Quantity, ExcelOrder.Instance)
            .ToList();

        return new StorageSelectionResult(sorted, matched, duplicates, blanks);
    }

    /// <summary>Порядок сортировки Excel по возрастанию: числа, затем текст, пустые в конце.</summary>
    internal sealed class ExcelOrder : IComparer<object?>
    {
        public static readonly ExcelOrder Instance = new();

        public int Compare(object? left, object? right)
        {
            var (leftRank, leftNumber, leftText) = Rank(left);
            var (rightRank, rightNumber, rightText) = Rank(right);

            if (leftRank != rightRank)
            {
                return leftRank.CompareTo(rightRank);
            }

            return leftRank switch
            {
                0 => leftNumber.CompareTo(rightNumber),
                1 => string.Compare(leftText, rightText, StringComparison.CurrentCultureIgnoreCase),
                _ => 0,
            };
        }

        private static (int Rank, double Number, string Text) Rank(object? value)
        {
            if (value is double number)
            {
                return (0, number, string.Empty);
            }

            var text = TextUtils.Normalize(TextUtils.CellToString(value));
            if (text.Length == 0)
            {
                return (2, 0d, string.Empty);
            }

            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
                ? (0, parsed, string.Empty)
                : (1, 0d, text);
        }
    }
}
