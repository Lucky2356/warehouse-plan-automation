using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Значения одной строки листа «Цены», собранные до обращения к Excel.</summary>
/// <param name="InvoiceIndex">
/// Номер листа инвойса, начиная с нуля. От него зависит, на какую колонку листа «для цен»
/// смотрят формулы «% логистики» и «Фактический курс».
/// </param>
/// <param name="Reference">Строка «Сводного прайса». null - штрихкода в прайсе нет.</param>
/// <param name="Markup">
/// Наценки по регламенту. На этапе подготовки всегда null: их ставят на пересчёте,
/// когда видно, сколько в поставке подгрупп.
/// </param>
public sealed record PriceRowValues(
    int Index,
    int InvoiceIndex,
    string Barcode,
    string Supply,
    double Units,
    double MaxPurchase,
    double PurchaseAmount,
    double PurchasePrice,
    PriceListRow? Reference,
    string Subgroup,
    MarkupRule? Markup,
    bool IsDuplicateBarcode);

/// <summary>Итог одного листа «для цен»: что записать в его колонку.</summary>
public sealed record SupplySummary(int InvoiceIndex, string Supply, double Quantity, double Amount);

/// <summary>Что нужно записать на листы «для цен» и «Цены».</summary>
public sealed class PriceSheetPlan
{
    public PriceSheetPlan(
        IReadOnlyList<PriceRowValues> rows,
        IReadOnlyList<SupplySummary> supplies,
        IReadOnlyList<ProcessingWarning> warnings)
    {
        Rows = rows;
        Supplies = supplies;
        Warnings = warnings;
    }

    public IReadOnlyList<PriceRowValues> Rows { get; }

    public IReadOnlyList<SupplySummary> Supplies { get; }

    public IReadOnlyList<ProcessingWarning> Warnings { get; }
}

/// <summary>
/// Сборка строк листа «Цены» из инвойса и справочников. Класс ничего не знает про Excel
/// и целиком покрывается тестами.
/// </summary>
public static class PriceRowBuilder
{
    public static PriceSheetPlan Build(
        IReadOnlyList<InvoiceSheet> invoices,
        IReadOnlyDictionary<string, PriceListRow> priceList,
        IReadOnlyDictionary<string, string> subgroups)
    {
        var rows = new List<PriceRowValues>();
        var supplies = new List<SupplySummary>();
        var warnings = new List<ProcessingWarning>();

        var barcodeCounts = CountBarcodes(invoices);
        var barcodePrices = SumPrices(invoices);
        var reportedDuplicates = new HashSet<string>(StringComparer.Ordinal);

        for (var invoiceIndex = 0; invoiceIndex < invoices.Count; invoiceIndex++)
        {
            var invoice = invoices[invoiceIndex];
            supplies.Add(new SupplySummary(
                invoiceIndex, invoice.SupplyNumber, invoice.TotalQuantity, invoice.TotalAmount));

            if (invoice.SupplyNumber.Length == 0)
            {
                warnings.Add(new ProcessingWarning(
                    "В шапке инвойса не найден номер поставки, колонка «Поставка» осталась пустой.",
                    invoice.SheetName,
                    invoice.SheetName));
            }

            foreach (var line in invoice.Lines)
            {
                priceList.TryGetValue(line.Barcode, out var reference);
                if (reference is null)
                {
                    warnings.Add(new ProcessingWarning(
                        "Штрихкод " + line.Barcode + " не найден в «" + PriceSchema.SummaryPriceSheet +
                        "»: код, модель, сектор и остальные справочные колонки не заполнены.",
                        invoice.SheetName + ", строка " + line.ExcelRow,
                        invoice.SheetName,
                        "A" + line.ExcelRow));
                }

                var subgroup = string.Empty;
                if (reference is not null && reference.Code.Length > 0 &&
                    !subgroups.TryGetValue(reference.Code, out subgroup!))
                {
                    subgroup = string.Empty;
                    warnings.Add(new ProcessingWarning(
                        "Код " + reference.Code + " не найден в «" + PriceSchema.DivisionPriceSheet +
                        "», подгруппа не заполнена.",
                        invoice.SheetName + ", строка " + line.ExcelRow,
                        invoice.SheetName,
                        "A" + line.ExcelRow));
                }

                // Повтор штрихкода: по инструкции в «цена_закупочная» пишется сумма «Мах_Закуп»
                // всех повторок - в каждую из них. Строка остаётся подсвеченной: повтор
                // в инвойсе стоит того, чтобы на него посмотреть.
                var duplicate = barcodeCounts[line.Barcode] > 1;
                if (duplicate && reportedDuplicates.Add(line.Barcode))
                {
                    warnings.Add(new ProcessingWarning(
                        "Штрихкод " + line.Barcode + " встречается в инвойсе " +
                        barcodeCounts[line.Barcode] + " раза: в «цена_закупочная» записана сумма " +
                        "«Мах_Закуп» этих строк - " +
                        barcodePrices[line.Barcode].ToString("0.####", System.Globalization.CultureInfo.CurrentCulture) +
                        ". В «Копия ШК» эти строки выделены одним цветом.",
                        PriceSchema.PricesSheet));
                }

                rows.Add(new PriceRowValues(
                    rows.Count,
                    invoiceIndex,
                    line.Barcode,
                    invoice.SupplyNumber,
                    line.Quantity,
                    line.Price,
                    line.Amount,
                    duplicate ? barcodePrices[line.Barcode] : line.Price,
                    reference,
                    subgroup ?? string.Empty,
                    null,
                    duplicate));
            }
        }

        if (rows.Count == 0)
        {
            warnings.Add(new ProcessingWarning(
                "В инвойсе не найдено ни одной строки товара - лист «Цены» не заполнен.",
                PriceSchema.InvoiceSheet));
        }

        return new PriceSheetPlan(rows, supplies, warnings);
    }

    /// <summary>
    /// Оставляет строки одного сектора: один файл распреда - один сектор, даже если
    /// в инвойсе их несколько. Строки без сектора остаются: их штрихкода нет в «Сводном
    /// прайсе», об этом уже написано отдельное замечание, и молча выбрасывать их нельзя.
    ///
    /// Итоги «для цен» пересчитываются по оставшимся строкам: в них должен попасть только
    /// тот товар, который остался в файле.
    /// </summary>
    public static PriceSheetPlan KeepSector(PriceSheetPlan plan, string sector)
    {
        var wanted = TextUtils.NormalizeKey(sector);
        if (wanted.Length == 0)
        {
            return plan;
        }

        bool Keep(PriceRowValues row)
        {
            var own = TextUtils.NormalizeKey(row.Reference?.Sector);
            return own.Length == 0 || string.Equals(own, wanted, StringComparison.Ordinal);
        }

        var dropped = plan.Rows.Where(row => !Keep(row)).ToList();
        if (dropped.Count == 0)
        {
            return plan;
        }

        var kept = plan.Rows.Where(Keep).ToList();

        // Поставки перенумеровываются подряд: если весь первый инвойс оказался чужого
        // сектора, оставшийся становится первым и смотрит на первую колонку «для цен».
        var order = kept
            .Select(row => row.InvoiceIndex)
            .Distinct()
            .OrderBy(index => index)
            .ToList();
        var moved = order.Select((old, position) => (old, position))
            .ToDictionary(pair => pair.old, pair => pair.position);

        var rows = kept
            .Select((row, position) => row with { Index = position, InvoiceIndex = moved[row.InvoiceIndex] })
            .ToList();

        var supplies = order
            .Select(old => new SupplySummary(
                moved[old],
                plan.Supplies.FirstOrDefault(supply => supply.InvoiceIndex == old)?.Supply ?? string.Empty,
                kept.Where(row => row.InvoiceIndex == old).Sum(row => row.Units),
                kept.Where(row => row.InvoiceIndex == old).Sum(row => row.PurchaseAmount)))
            .ToList();

        var warnings = plan.Warnings.ToList();
        warnings.Add(new ProcessingWarning(
            "В инвойсе несколько секторов. По названию файла оставлен «" + TextUtils.Normalize(sector) +
            "», не попали в файл: " +
            string.Join(", ", dropped
                .GroupBy(row => TextUtils.Normalize(row.Reference?.Sector), StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(group => group.Count())
                .Select(group => "«" + group.Key + "» - " + group.Count() + " шт.")) +
            ". Для них нужен отдельный файл.",
            PriceSchema.PricesSheet));

        return new PriceSheetPlan(rows, supplies, warnings);
    }

    private static Dictionary<string, int> CountBarcodes(IReadOnlyList<InvoiceSheet> invoices)
    {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var line in invoices.SelectMany(invoice => invoice.Lines))
        {
            counts[line.Barcode] = counts.TryGetValue(line.Barcode, out var current) ? current + 1 : 1;
        }

        return counts;
    }

    private static Dictionary<string, double> SumPrices(IReadOnlyList<InvoiceSheet> invoices)
    {
        var sums = new Dictionary<string, double>(StringComparer.Ordinal);
        foreach (var line in invoices.SelectMany(invoice => invoice.Lines))
        {
            sums[line.Barcode] = (sums.TryGetValue(line.Barcode, out var current) ? current : 0d) + line.Price;
        }

        return sums;
    }
}
