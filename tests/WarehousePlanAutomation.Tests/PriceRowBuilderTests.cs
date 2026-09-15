using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class PriceRowBuilderTests
{
    private static PriceListRow Reference(string barcode, string code) => new(
        barcode, code, "SLEEP MASK_FW25-26_3", "МЕЛКИЕ АКСЕССУАРЫ", "РАЗНОЕ",
        "МАСКА ДЛЯ СНА", "431-2802", "СЕРЫЙ-ЧЕРНЫЙ", "WINTER 26-27", "TU");

    private static InvoiceSheet Invoice(string supply, params InvoiceLine[] lines) =>
        new("Invoice", supply, lines);

    private static PriceSheetPlan Build(
        IReadOnlyList<InvoiceSheet> invoices,
        Dictionary<string, PriceListRow>? priceList = null,
        Dictionary<string, string>? subgroups = null) =>
        PriceRowBuilder.Build(
            invoices,
            priceList ?? new Dictionary<string, PriceListRow>(StringComparer.Ordinal),
            subgroups ?? new Dictionary<string, string>(StringComparer.Ordinal));

    [Fact]
    public void СтрокиСобираютсяИзИнвойсаИСправочников()
    {
        var plan = Build(
            new[] { Invoice("431-329", new InvoiceLine(25, "1431280202", 1010d, 6.55d, 6615.5d)) },
            new Dictionary<string, PriceListRow>(StringComparer.Ordinal)
            {
                ["1431280202"] = Reference("1431280202", "155588"),
            },
            new Dictionary<string, string>(StringComparer.Ordinal) { ["155588"] = "МАСКА ДЛЯ СНА/" });

        var row = Assert.Single(plan.Rows);
        Assert.Equal("431-329", row.Supply);
        Assert.Equal(1010d, row.Units);
        Assert.Equal(6.55d, row.MaxPurchase);
        Assert.Equal(6615.5d, row.PurchaseAmount);
        Assert.Equal("155588", row.Reference!.Code);
        Assert.Equal("МАСКА ДЛЯ СНА/", row.Subgroup);
        Assert.Null(row.Markup);
        Assert.Empty(plan.Warnings);
    }

    [Fact]
    public void ЦенаЗакупочнаяРавнаМахЗакуп()
    {
        var plan = Build(new[] { Invoice("431-329", new InvoiceLine(25, "1", 10d, 6.55d, 65.5d)) });

        Assert.Equal(6.55d, plan.Rows[0].PurchasePrice);
    }

    [Fact]
    public void ПовторШтрихкода_ЦенаЗакупочнаяСуммаПовторокИСтрокаПомечена()
    {
        // Инструкция: при повторе «Копия ШК» в «цена_закупочная» пишется сумма «Мах_Закуп» повторок.
        var plan = Build(new[]
        {
            Invoice(
                "431-329",
                new InvoiceLine(25, "1431280202", 500d, 6.55d, 3275d),
                new InvoiceLine(26, "1431280202", 510d, 3.20d, 1632d),
                new InvoiceLine(27, "1431333901", 840d, 6.55d, 5502d)),
        });

        Assert.True(plan.Rows[0].IsDuplicateBarcode);
        Assert.True(plan.Rows[1].IsDuplicateBarcode);
        Assert.False(plan.Rows[2].IsDuplicateBarcode);
        Assert.Equal(9.75d, plan.Rows[0].PurchasePrice, 6);
        Assert.Equal(9.75d, plan.Rows[1].PurchasePrice, 6);
        Assert.Equal(6.55d, plan.Rows[2].PurchasePrice);

        // «Мах_Закуп» у каждой строки остаётся своим.
        Assert.Equal(6.55d, plan.Rows[0].MaxPurchase);
        Assert.Equal(3.20d, plan.Rows[1].MaxPurchase);
        Assert.Single(plan.Warnings, w => w.Message.Contains("встречается в инвойсе"));
    }

    [Fact]
    public void ШтрихкодНеНайденВПрайсе_ДаётЗамечаниеИПустуюСправку()
    {
        var plan = Build(new[] { Invoice("431-329", new InvoiceLine(25, "1431280202", 10d, 1d, 10d)) });

        Assert.Null(plan.Rows[0].Reference);
        Assert.Equal(string.Empty, plan.Rows[0].Subgroup);
        Assert.Contains(plan.Warnings, w => w.Message.Contains("1431280202"));
    }

    [Fact]
    public void ЗамечаниеПоСтрокеИнвойсаЗнаетСвоёМесто()
    {
        // По такому замечанию окно открывает книгу сразу на нужной строке инвойса,
        // поэтому лист и ячейка нужны не текстом, а отдельно.
        var plan = Build(new[] { Invoice("431-329", new InvoiceLine(25, "1431280202", 10d, 1d, 10d)) });

        var warning = Assert.Single(plan.Warnings, w => w.Message.Contains("1431280202"));

        Assert.True(warning.CanNavigate);
        Assert.Equal("Invoice", warning.Sheet);
        Assert.Equal("A25", warning.Cell);
    }

    [Fact]
    public void ОбщееЗамечаниеПереходаНеПредлагает()
    {
        // «Регламент наценок» - это про правило, а не про ячейку: вести туда некуда.
        var warning = new ProcessingWarning("Наценки не проставлены.", "Регламент наценок");

        Assert.False(warning.CanNavigate);
    }

    [Fact]
    public void КодНеНайденВПрайсеПоПодразделениям_ПодгруппаПустаяИЕстьЗамечание()
    {
        var plan = Build(
            new[] { Invoice("431-329", new InvoiceLine(25, "1431280202", 10d, 1d, 10d)) },
            new Dictionary<string, PriceListRow>(StringComparer.Ordinal)
            {
                ["1431280202"] = Reference("1431280202", "155588"),
            });

        Assert.Equal(string.Empty, plan.Rows[0].Subgroup);
        Assert.Contains(plan.Warnings, w => w.Message.Contains("155588"));
    }

    [Fact]
    public void ДваЛистаИнвойса_ДаютДваИтогаИРазныеНомераЛистов()
    {
        var plan = Build(new[]
        {
            Invoice("431-329", new InvoiceLine(25, "1", 100d, 2d, 200d)),
            Invoice("431-332", new InvoiceLine(25, "2", 50d, 4d, 200d)),
        });

        Assert.Equal(2, plan.Supplies.Count);
        Assert.Equal("431-329", plan.Supplies[0].Supply);
        Assert.Equal(100d, plan.Supplies[0].Quantity);
        Assert.Equal("431-332", plan.Supplies[1].Supply);
        Assert.Equal(200d, plan.Supplies[1].Amount);

        Assert.Equal(0, plan.Rows[0].InvoiceIndex);
        Assert.Equal(1, plan.Rows[1].InvoiceIndex);
    }

    [Fact]
    public void ПустойНомерПоставки_ДаётЗамечание()
    {
        var plan = Build(new[] { Invoice(string.Empty, new InvoiceLine(25, "1", 10d, 1d, 10d)) });

        Assert.Contains(plan.Warnings, w => w.Message.Contains("номер поставки"));
    }

    [Fact]
    public void ПустойИнвойс_ДаётЗамечание()
    {
        var plan = Build(new[] { Invoice("431-329") });

        Assert.Empty(plan.Rows);
        Assert.Contains(plan.Warnings, w => w.Message.Contains("ни одной строки товара"));
    }

    // ===== Один файл - один сектор =====

    private static PriceListRow OfSector(string barcode, string sector) => new(
        barcode, "155588", "модель", sector, "РАЗНОЕ", "имя", "арт", "цвет", "сезон", "TU");

    private static PriceSheetPlan TwoSectors() => Build(
        new[]
        {
            Invoice(
                "431-329",
                new InvoiceLine(25, "1", 10d, 2d, 20d),
                new InvoiceLine(26, "2", 30d, 3d, 90d),
                new InvoiceLine(27, "3", 5d, 1d, 5d)),
        },
        new Dictionary<string, PriceListRow>(StringComparer.Ordinal)
        {
            ["1"] = OfSector("1", "БИЖУТЕРИЯ"),
            ["2"] = OfSector("2", "УКРАШЕНИЯ ДЛЯ ВОЛОС"),
        });

    [Fact]
    public void ЛишнийСекторНеПопадаетВФайл()
    {
        var plan = PriceRowBuilder.KeepSector(TwoSectors(), "БИЖУТЕРИЯ");

        // Осталась бижутерия и строка без сектора: её штрихкода нет в прайсе,
        // и выбрасывать её молча нельзя.
        Assert.Equal(new[] { "1", "3" }, plan.Rows.Select(row => row.Barcode));
        Assert.Equal(new[] { 0, 1 }, plan.Rows.Select(row => row.Index));
        Assert.Contains(plan.Warnings, w => w.Message.Contains("УКРАШЕНИЯ ДЛЯ ВОЛОС"));
    }

    [Fact]
    public void ИтогиПоставкиСчитаютсяПоОставшимсяСтрокам()
    {
        var supply = Assert.Single(PriceRowBuilder.KeepSector(TwoSectors(), "БИЖУТЕРИЯ").Supplies);

        Assert.Equal("431-329", supply.Supply);
        Assert.Equal(15d, supply.Quantity);
        Assert.Equal(25d, supply.Amount);
    }

    [Fact]
    public void ПоставкиПеренумеровываютсяПодряд()
    {
        var plan = Build(
            new[]
            {
                Invoice("431-329", new InvoiceLine(25, "1", 10d, 2d, 20d)),
                Invoice("431-332", new InvoiceLine(25, "2", 30d, 3d, 90d)),
            },
            new Dictionary<string, PriceListRow>(StringComparer.Ordinal)
            {
                ["1"] = OfSector("1", "СУМКИ"),
                ["2"] = OfSector("2", "БИЖУТЕРИЯ"),
            });

        var kept = PriceRowBuilder.KeepSector(plan, "БИЖУТЕРИЯ");

        // От первого инвойса не осталось ничего: второй становится первой колонкой «для цен».
        Assert.Equal(0, Assert.Single(kept.Rows).InvoiceIndex);
        Assert.Equal(0, Assert.Single(kept.Supplies).InvoiceIndex);
        Assert.Equal("431-332", kept.Supplies[0].Supply);
    }

    [Fact]
    public void ОдинСектор_ПланНеМеняется()
    {
        var plan = Build(
            new[] { Invoice("431-329", new InvoiceLine(25, "1", 10d, 2d, 20d)) },
            new Dictionary<string, PriceListRow>(StringComparer.Ordinal)
            {
                ["1"] = OfSector("1", "БИЖУТЕРИЯ"),
            });

        Assert.Same(plan, PriceRowBuilder.KeepSector(plan, "БИЖУТЕРИЯ"));
    }
}
