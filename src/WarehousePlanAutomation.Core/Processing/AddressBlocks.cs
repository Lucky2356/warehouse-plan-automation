using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Строка «А2, А3» или «МП», у которой «Учитывать» больше нуля.</summary>
public sealed record CountedRow(
    object? Address,
    object? Container,
    object? Code,
    object? AddressNumber,
    object? ContainerNumber);

/// <summary>Строка листа «адреса». <paramref name="Joined"/> заполнен у первой строки кода.</summary>
public sealed record AddressLine(object? Address, object? Code, object? AddressNumber, string Joined);

/// <summary>Строка листа «тары». <paramref name="Joined"/> заполнен у первой строки кода.</summary>
public sealed record ContainerLine(
    object? Address,
    object? Container,
    object? Code,
    object? AddressNumber,
    object? ContainerNumber,
    string Joined);

/// <summary>
/// Листы «адреса» и «тары»: по каждому коду - все адреса и все тары одной строкой,
/// чтобы «итог» мог подтянуть их ВПР по коду.
///
/// Вручную это ОБЪЕДИНИТЬ, растянутое на строки одного кода, в первой строке кода,
/// а у остальных строк кода ячейка очищается. Строки кода идут подряд: «А2, А3»
/// отсортирован по коду. На «адресах» повтор адреса у одного кода убирается.
/// </summary>
public static class AddressBlocks
{
    private const string Separator = ", ";

    public static IReadOnlyList<AddressLine> Addresses(IReadOnlyList<CountedRow> rows)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var unique = rows
            .Where(row => seen.Add(Text(row.Address) + "|" + ReceivingStockRules.CodeKey(row.Code)))
            .ToList();

        var lines = new List<AddressLine>(unique.Count);
        foreach (var group in Groups(unique))
        {
            var joined = Join(group.Select(row => row.Address));
            for (var i = 0; i < group.Count; i++)
            {
                var row = group[i];
                lines.Add(new AddressLine(row.Address, row.Code, row.AddressNumber, i == 0 ? joined : string.Empty));
            }
        }

        return lines;
    }

    public static IReadOnlyList<ContainerLine> Containers(IReadOnlyList<CountedRow> rows)
    {
        var lines = new List<ContainerLine>(rows.Count);
        foreach (var group in Groups(rows))
        {
            var joined = Join(group.Select(row => row.Container));
            for (var i = 0; i < group.Count; i++)
            {
                var row = group[i];
                lines.Add(new ContainerLine(
                    row.Address, row.Container, row.Code, row.AddressNumber, row.ContainerNumber,
                    i == 0 ? joined : string.Empty));
            }
        }

        return lines;
    }

    /// <summary>Подряд идущие строки одного кода.</summary>
    private static IEnumerable<List<CountedRow>> Groups(IReadOnlyList<CountedRow> rows)
    {
        var current = new List<CountedRow>();
        var currentKey = string.Empty;

        foreach (var row in rows)
        {
            var key = ReceivingStockRules.CodeKey(row.Code);
            if (current.Count > 0 && !string.Equals(key, currentKey, StringComparison.Ordinal))
            {
                yield return current;
                current = new List<CountedRow>();
            }

            currentKey = key;
            current.Add(row);
        }

        if (current.Count > 0)
        {
            yield return current;
        }
    }

    /// <summary>Как ОБЪЕДИНИТЬ(", "; 1; …): пустые значения пропускаются.</summary>
    private static string Join(IEnumerable<object?> values) =>
        string.Join(Separator, values.Select(Text).Where(text => text.Length > 0));

    private static string Text(object? value) => TextUtils.Normalize(TextUtils.CellToString(value));
}
