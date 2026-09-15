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
/// Сопоставление идёт в четыре прохода: полное совпадение, полное совпадение без
/// приписанной единицы измерения («цена, CNY»), начало строки, вхождение.
/// Это позволяет находить и «Дата в сети (без целевой даты...)», и «Разница ед».
/// </summary>
public static class HeaderResolver
{
    /// <summary>Сколько первых строк листа просматривается в поиске строки заголовков.</summary>
    public const int DefaultScanRows = 15;

    /// <summary>После этих символов в заголовке идёт единица измерения, а не продолжение названия.</summary>
    private static readonly char[] UnitSeparators = { ',', '(', '[' };

    private const string CurrencySigns = "¥$€₽£";

    public static HeaderMap Resolve(
        SheetGrid grid,
        string sheetName,
        IReadOnlyList<ColumnSpec> specs,
        int scanRows = DefaultScanRows)
    {
        var bestMissing = specs.Select(s => s.DisplayName).ToList();
        var bestRow = grid.FirstRow;

        var lastScanned = Math.Min(grid.LastRow, grid.FirstRow + scanRows - 1);
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
        if (byExact >= 0)
        {
            return byExact;
        }

        var withoutUnit = Scan(
            headers, used, (header, alias) => string.Equals(WithoutUnit(header), alias, StringComparison.Ordinal));
        var byWithoutUnit = withoutUnit(spec);
        if (byWithoutUnit >= 0 || spec.ExactOnly)
        {
            return byWithoutUnit;
        }

        var prefix = Scan(headers, used, (header, alias) => header.StartsWith(alias, StringComparison.Ordinal));
        var byPrefix = prefix(spec);
        if (byPrefix >= 0 || spec.PrefixOnly)
        {
            return byPrefix;
        }

        var contains = Scan(headers, used, (header, alias) => header.Contains(alias, StringComparison.Ordinal));
        return contains(spec);
    }

    /// <summary>
    /// Отбрасывает единицу измерения, приписанную к названию колонки: «цена, CNY»,
    /// «сумма, CNY», «количество, шт.», «вес (кг)», «цена CN¥». То, что осталось,
    /// сравнивается полным совпадением - поэтому «цена продажи» колонкой «цена» не станет.
    /// </summary>
    private static string WithoutUnit(string header)
    {
        var cut = header.IndexOfAny(UnitSeparators);
        if (cut > 0)
        {
            return header[..cut].TrimEnd();
        }

        var space = header.LastIndexOf(' ');
        return space > 0 && IsCurrency(header[(space + 1)..]) ? header[..space] : header;
    }

    /// <summary>
    /// Без запятой хвост отбрасывается только если это обозначение валюты: «CNY», «CN¥», «$».
    /// Так «цена CNY» находится, а «цена продажи» и «% выполнения» остаются нетронутыми.
    /// </summary>
    private static bool IsCurrency(string tail)
    {
        var value = tail.TrimEnd('.');
        if (value.Length is 0 or > 5)
        {
            return false;
        }

        foreach (var symbol in value)
        {
            if (symbol is not (>= 'a' and <= 'z') && !CurrencySigns.Contains(symbol))
            {
                return false;
            }
        }

        return true;
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
