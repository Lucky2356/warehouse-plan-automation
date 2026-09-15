using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Одна строка выгрузки «Остатки Н», отобранная по сектору.</summary>
public sealed record StockSourceRow(string Acr, IReadOnlyList<object?> Values);

/// <summary>
/// Содержимое листа «остатки»: подписи строк и значения по каждому АЦР.
/// Лист - это транспонированная выгрузка, поэтому строка здесь соответствует
/// колонке «Остатки Н», а колонка - её строке.
/// </summary>
public sealed class StockSheetContent
{
    public StockSheetContent(IReadOnlyList<string> labels, IReadOnlyList<StockSourceRow> rows)
    {
        Labels = labels;
        Rows = rows;
    }

    /// <summary>Подписи строк листа «остатки» - они же заголовки колонок «Остатки Н».</summary>
    public IReadOnlyList<string> Labels { get; }

    /// <summary>Отобранные строки выгрузки. Каждая станет колонкой листа «остатки».</summary>
    public IReadOnlyList<StockSourceRow> Rows { get; }

    public int RowCount => Labels.Count;

    public int ColumnCount => Rows.Count;

    /// <summary>
    /// Строки магазинов, по которым «Распред» считает остаток: «001 тз» и «001 в_пути».
    /// Строка «001 прод» в остаток не входит - проданное остатком не является,
    /// и в разобранном файле у неё код магазина не проставлен.
    /// </summary>
    public bool IsStockRow(int index)
    {
        var label = TextUtils.NormalizeKey(Labels[index]);
        return PriceSchema.StockSource.StockRowSuffixes.Any(suffix =>
            label.EndsWith(" " + suffix, StringComparison.Ordinal));
    }
}

/// <summary>
/// Сборка листа «остатки» из выгрузки «Остатки Н».
///
/// Саму выгрузку программа не меняет: «АЦР» собирается из артикула, цвета и размера
/// в памяти, лишние колонки в памяти же отбрасываются. Так лист остаётся тем, что
/// выгрузили, и повторный запуск не зависит от того, правили его руками или нет.
/// </summary>
public static class StockSheetBuilder
{
    /// <summary>Колонки, попадающие в «остатки», в порядке листа: «АЦР» первой.</summary>
    public static IReadOnlyList<int> ChooseColumns(SheetGrid grid, int headerRow, out IReadOnlyList<string> labels)
    {
        var chosen = new List<int>();
        var names = new List<string>();

        // «АЦР» идёт первой колонкой и собирается сама, если её в выгрузке нет.
        chosen.Add(0);
        names.Add(PriceSchema.StockSource.Acr);

        for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
        {
            var header = TextUtils.Normalize(grid.Text(headerRow, column));
            var key = TextUtils.NormalizeKey(header);

            if (header.Length == 0 ||
                key == TextUtils.NormalizeKey(PriceSchema.StockSource.Acr) ||
                PriceSchema.StockSource.Dropped.Contains(key, StringComparer.Ordinal))
            {
                continue;
            }

            chosen.Add(column);
            names.Add(header);
        }

        labels = names;
        return chosen;
    }

    /// <summary>«АЦР» - это артикул, цвет и размер подряд, без разделителей.</summary>
    public static string BuildAcr(string? article, string? color, string? size) =>
        TextUtils.Normalize(article) + TextUtils.Normalize(color) + TextUtils.Normalize(size);
}
