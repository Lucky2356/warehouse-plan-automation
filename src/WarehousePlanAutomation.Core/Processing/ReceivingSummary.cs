using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Суммы по коду, из которых складываются колонки «А2, А3» … «Поставки МП» листа «итог».</summary>
public sealed record SummarySources(
    IReadOnlyDictionary<string, double> Storage,
    IReadOnlyDictionary<string, double> Marketplace,
    IReadOnlyDictionary<string, double> Collected,
    IReadOnlyDictionary<string, double> NotCollected,
    IReadOnlyDictionary<string, double> MarketplaceSupplies);

/// <summary>Строка «итога» перед вторым этапом - то, что нужно для решения.</summary>
/// <param name="Restock">«Допоставить» как есть: пусто и ноль - разные вещи.</param>
public sealed record SummaryState(
    int Index,
    object? Restock,
    object? Quantity,
    object? QuantityMarketplace,
    object? Collected);

/// <summary>Откуда «итог» берёт место хранения и тару.</summary>
public enum PlaceSource
{
    None,

    /// <summary>С листов «адреса» и «тары»: товар набирается на хранении.</summary>
    StorageAddresses,

    /// <summary>С листа «Поставки собраны»: товар берётся из собранной поставки.</summary>
    CollectedSupply,
}

/// <summary>
/// Что второй этап делает со строкой «итога».
/// <paramref name="RestockFromMarketplace"/> - что вписать в пустое «Допоставить».
/// </summary>
public sealed record SummaryFill(int Index, PlaceSource Place, double? RestockFromMarketplace);

/// <summary>
/// Лист «итог».
///
/// На первом этапе в него попадают строки «Остатков», по коду которых есть хоть что-то
/// в «А2, А3», «МП» или в одной из трёх поставок: вручную это фильтр «ноль и меньше»
/// сразу по пяти колонкам и удаление отфильтрованного.
///
/// На втором этапе, когда «Допоставить» уже решено, заполняются место хранения и тара,
/// а пустое «Допоставить» получает «Количество МП».
/// </summary>
public static class ReceivingSummary
{
    /// <summary>Номера строк, которые остаются на «итоге».</summary>
    public static IReadOnlyList<int> SelectRows(IReadOnlyList<object?> codes, SummarySources sources)
    {
        var kept = new List<int>();
        for (var i = 0; i < codes.Count; i++)
        {
            var key = ReceivingStockRules.CodeKey(codes[i]);
            if (key.Length == 0)
            {
                continue;
            }

            if (Positive(sources.Storage, key) || Positive(sources.Marketplace, key) ||
                Positive(sources.Collected, key) || Positive(sources.NotCollected, key) ||
                Positive(sources.MarketplaceSupplies, key))
            {
                kept.Add(i);
            }
        }

        return kept;
    }

    /// <summary>
    /// «Ответ МП» по коду - так, как его находит ВПР: берётся первая строка кода,
    /// и в «Количество МП» идёт только то, что больше нуля.
    /// </summary>
    public static Dictionary<string, double> MarketplaceAnswers(IEnumerable<(object? Code, object? Answer)> rows)
    {
        var first = new Dictionary<string, double?>(StringComparer.Ordinal);
        foreach (var (code, answer) in rows)
        {
            var key = ReceivingStockRules.CodeKey(code);
            if (key.Length == 0 || first.ContainsKey(key))
            {
                continue;
            }

            first[key] = CellError.IsError(answer) ? null : TextUtils.CellToDouble(answer);
        }

        return first
            .Where(pair => pair.Value is > 0)
            .ToDictionary(pair => pair.Key, pair => pair.Value!.Value, StringComparer.Ordinal);
    }

    /// <summary>
    /// Решения второго этапа, в порядке инструкции: место хранения с адресов для строк
    /// с «Количеством» больше нуля; из собранной поставки - для строк, где «Допоставить»
    /// заполнено и «Поставки собраны» больше нуля (это перекрывает адреса); «Количество МП» -
    /// в пустое «Допоставить». Заполненность «Допоставить» смотрится до того, как в него
    /// что-то вписано.
    /// </summary>
    public static IReadOnlyList<SummaryFill> PlanFills(IReadOnlyList<SummaryState> rows)
    {
        var fills = new List<SummaryFill>();

        foreach (var row in rows)
        {
            var restockFilled = TextUtils.Normalize(TextUtils.CellToString(row.Restock)).Length > 0;

            var place = PlaceSource.None;
            if (Number(row.Quantity) > 0)
            {
                place = PlaceSource.StorageAddresses;
            }

            if (restockFilled && Number(row.Collected) > 0)
            {
                place = PlaceSource.CollectedSupply;
            }

            var fromMarketplace = !restockFilled && Number(row.QuantityMarketplace) is > 0 and var mp
                ? mp
                : (double?)null;

            if (place != PlaceSource.None || fromMarketplace is not null)
            {
                fills.Add(new SummaryFill(row.Index, place, fromMarketplace));
            }
        }

        return fills;
    }

    private static bool Positive(IReadOnlyDictionary<string, double> sums, string key) =>
        sums.TryGetValue(key, out var value) && value > 0;

    private static double Number(object? value) =>
        CellError.IsError(value) ? 0d : TextUtils.CellToDouble(value) ?? 0d;
}
