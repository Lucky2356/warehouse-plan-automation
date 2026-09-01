using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>Результат поиска строки заголовков и позиций колонок.</summary>
public sealed class HeaderMap
{
    private readonly Dictionary<string, int> _columns;

    public HeaderMap(int headerRow, Dictionary<string, int> columns)
    {
        HeaderRow = headerRow;
        _columns = columns;
    }

    public int HeaderRow { get; }

    public IReadOnlyDictionary<string, int> Columns => _columns;

    public int this[string displayName] => _columns[displayName];

    public bool TryGet(string displayName, out int column) => _columns.TryGetValue(displayName, out column);
}

/// <summary>
/// Поиск строки заголовков и колонок по нормализованным названиям.
/// Сопоставление идёт в три прохода: полное совпадение, начало строки, вхождение.
/// Это позволяет находить и «Дата в сети (без целевой даты...)», и «Разница ед».
/// </summary>
public static class HeaderResolver
{
    private const int MaxHeaderRowScan = 15;

    public static HeaderMap Resolve(SheetGrid grid, string sheetName, IReadOnlyList<ColumnSpec> specs)
    {
        var bestMissing = specs.Select(s => s.DisplayName).ToList();
        var bestRow = grid.FirstRow;

        var lastScanned = Math.Min(grid.LastRow, grid.FirstRow + MaxHeaderRowScan - 1);
        for (var row = grid.FirstRow; row <= lastScanned; row++)
        {
            var resolved = TryResolveRow(grid, row, specs, out var missing);
            if (missing.Count == 0)
            {
                return new HeaderMap(row, resolved);
            }

            if (missing.Count < bestMissing.Count)
            {
                bestMissing = missing;
                bestRow = row;
            }
        }

        var problems = bestMissing
            .Select(name => "на листе «" + sheetName + "» не найдена колонка «" + name +
                            "» (наиболее похожая строка заголовков: " + bestRow + ")")
            .ToList();
        throw new WorkbookValidationException(problems);
    }

    private static Dictionary<string, int> TryResolveRow(
        SheetGrid grid,
        int row,
        IReadOnlyList<ColumnSpec> specs,
        out List<string> missing)
    {
        var headers = new string[grid.ColumnCount];
        for (var i = 0; i < grid.ColumnCount; i++)
        {
            headers[i] = TextUtils.NormalizeKey(grid.Text(row, grid.FirstColumn + i));
        }

        var result = new Dictionary<string, int>(StringComparer.Ordinal);
        missing = new List<string>();
        var used = new HashSet<int>();

        foreach (var spec in specs)
        {
            var column = Match(headers, spec, used);
            if (column < 0)
            {
                missing.Add(spec.DisplayName);
                continue;
            }

            used.Add(column);
            result[spec.DisplayName] = grid.FirstColumn + column;
        }

        return result;
    }

    private static int Match(string[] headers, ColumnSpec spec, HashSet<int> used)
    {
        var exact = Scan(headers, used, (header, alias) => string.Equals(header, alias, StringComparison.Ordinal));
        var byExact = exact(spec);
        if (byExact >= 0 || spec.ExactOnly)
        {
            return byExact;
        }

        var prefix = Scan(headers, used, (header, alias) => header.StartsWith(alias, StringComparison.Ordinal));
        var byPrefix = prefix(spec);
        if (byPrefix >= 0)
        {
            return byPrefix;
        }

        var contains = Scan(headers, used, (header, alias) => header.Contains(alias, StringComparison.Ordinal));
        return contains(spec);
    }

    private static Func<ColumnSpec, int> Scan(string[] headers, HashSet<int> used, Func<string, string, bool> predicate)
    {
        return spec =>
        {
            for (var i = 0; i < headers.Length; i++)
            {
                if (used.Contains(i) || headers[i].Length == 0)
                {
                    continue;
                }

                foreach (var alias in spec.Aliases)
                {
                    if (predicate(headers[i], alias))
                    {
                        return i;
                    }
                }
            }

            return -1;
        };
    }
}
