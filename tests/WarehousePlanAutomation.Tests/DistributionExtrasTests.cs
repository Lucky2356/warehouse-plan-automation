using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class AnalogueLookupTests
{
    [Fact]
    public void АЦПоставкиСправа_АналогомСтановитсяАЦСлева()
    {
        var lookup = AnalogueLookup.Build(
            new[]
            {
                new AnaloguePair("2416-026БЕЛЫЙ-ЧЕРНЫЙ", "2416-045БЕЛЫЙ-ЧЕРНЫЙ"),
                new AnaloguePair("2405-001ПРОЗРАЧНЫЙ", "2451-015ПРОЗРАЧНЫЙ"),
            },
            new[] { "2416-045белый-черный", "431-2802СЕРЫЙ-ЧЕРНЫЙ" });

        Assert.Equal("2416-026БЕЛЫЙ-ЧЕРНЫЙ", AnalogueLookup.ValueFor(lookup, "2416-045БЕЛЫЙ-ЧЕРНЫЙ"));
        Assert.Equal(AnalogueLookup.None, AnalogueLookup.ValueFor(lookup, "431-2802СЕРЫЙ-ЧЕРНЫЙ"));
    }

    [Fact]
    public void АЦПоставкиСлева_ЗелёнаяСтрока_АналогомНеСчитается()
    {
        // Строка, где АЦ поставки стоит в первой колонке, при ручной разметке зелёная
        // и в колонки J:K не переносится.
        var lookup = AnalogueLookup.Build(
            new[] { new AnaloguePair("2416-045БЕЛЫЙ-ЧЕРНЫЙ", "2416-099БЕЛЫЙ-ЧЕРНЫЙ") },
            new[] { "2416-045БЕЛЫЙ-ЧЕРНЫЙ" });

        Assert.Equal(AnalogueLookup.None, AnalogueLookup.ValueFor(lookup, "2416-045БЕЛЫЙ-ЧЕРНЫЙ"));
    }

    [Fact]
    public void ОбаАЦВПоставке_АналогомНеСчитается()
    {
        var lookup = AnalogueLookup.Build(
            new[] { new AnaloguePair("2412-010ШОКОЛАДНЫЙ", "2412-018ШОКОЛАДНЫЙ") },
            new[] { "2412-010ШОКОЛАДНЫЙ", "2412-018ШОКОЛАДНЫЙ" });

        Assert.Empty(lookup);
    }

    [Fact]
    public void НесколькоСтрокСОднимАЦ_БерётсяПерваяКакУВПР()
    {
        var lookup = AnalogueLookup.Build(
            new[]
            {
                new AnaloguePair("1-001КРАСНЫЙ", "9-001КРАСНЫЙ"),
                new AnaloguePair("2-001КРАСНЫЙ", "9-001КРАСНЫЙ"),
            },
            new[] { "9-001КРАСНЫЙ" });

        Assert.Equal("1-001КРАСНЫЙ", AnalogueLookup.ValueFor(lookup, "9-001КРАСНЫЙ"));
    }
}

public class ApprovedPriceFilesTests
{
    private static ApprovedPriceFile File(string name, int day = 1) =>
        new(@"\\fs\КМ\Согласование цен\FW 26-27\" + name, new DateTime(2026, 9, day));

    [Theory]
    [InlineData("431-329", "431", 329)]
    [InlineData(" C2518-084 ", "С2518", 84)]
    [InlineData("С2518-084", "С2518", 84)]
    public void НомерПоставкиРазбирается(string supply, string prefix, int number)
    {
        var key = ApprovedPriceFiles.Parse(supply);

        Assert.Equal(new SupplyKey(prefix, number), key);
    }

    [Fact]
    public void СначалаФайлПоставки_ПотомБлижайшиеПоНомеру()
    {
        var order = ApprovedPriceFiles.Order(
            "431-329",
            new[]
            {
                File("431-335 Маски.xlsx"),
                File("431-329 Маски для сна.xlsx"),
                File("431-328 Бижутерия.xlsx"),
                File("318-329 Носки.xlsx"),
                File("431-3290 Другое.xlsx"),
                File("431-331.xlsx"),
            });

        Assert.Equal(new[] { "431-329 Маски для сна.xlsx" }, order.Exact.Select(f => Path.GetFileName(f.Path)));
        Assert.Equal(
            new[] { "431-328 Бижутерия.xlsx", "431-331.xlsx", "431-335 Маски.xlsx", "431-3290 Другое.xlsx" },
            order.Nearest.Select(f => Path.GetFileName(f.Path)));
    }

    [Fact]
    public void ЛатинскаяИКириллическаяБукваВНомереРавны()
    {
        var order = ApprovedPriceFiles.Order("C2518-084", new[] { File("С2518-084 Бижутерия.xlsx") });

        Assert.Single(order.Exact);
    }

    [Fact]
    public void НесколькоФайловПоставки_НовыйПервым()
    {
        var order = ApprovedPriceFiles.Order(
            "431-329",
            new[] { File("431-329 старый.xlsx", 2), File("431-329 новый.xlsx", 10) });

        Assert.Equal("431-329 новый.xlsx", Path.GetFileName(order.Exact[0].Path));
    }

    [Theory]
    [InlineData("431-329.xlsx", true)]
    [InlineData("431-329.xls", true)]
    [InlineData("~$431-329.xlsx", false)]
    [InlineData("431-329.pdf", false)]
    public void ВременныеИЧужиеФайлыПропускаются(string name, bool expected)
    {
        Assert.Equal(expected, ApprovedPriceFiles.IsWorkbook(name));
    }
}

public class ApprovedPriceSheetReaderTests
{
    [Fact]
    public void ТаблицаНаходитсяПоШКИРозничнойЦене()
    {
        var grid = SheetGrid.FromRows(1, 1, new[]
        {
            new object?[] { "Согласование цен", null, null, null },
            new object?[] { null, null, null, null },
            new object?[] { "ШК", "Артикул", "Предложение УТЗ", "Розничная цена" },
            new object?[] { 1431280202d, "431-2802", 649d, 599d },
            new object?[] { "1431333901", "431-3339", 649d, CellError.NotAvailable },
            new object?[] { 1431280202d, "повтор", 1d, 1d },
        });

        var table = ApprovedPriceSheetReader.FindTable(grid);

        Assert.NotNull(table);
        Assert.Equal(3, table!.HeaderRow);
        Assert.Equal(1, table.BarcodeColumn);
        Assert.Equal(4, table.PriceColumn);

        var prices = new Dictionary<string, double>(StringComparer.Ordinal);
        var added = ApprovedPriceSheetReader.Read(grid, table.BarcodeColumn, grid, table.PriceColumn, prices);

        // Ошибка в цене не цена, повтор штрихкода не перезаписывает первый - как у ВПР.
        Assert.Equal(1, added);
        Assert.Equal(599d, prices["1431280202"]);
        Assert.False(prices.ContainsKey("1431333901"));
    }

    [Fact]
    public void БезРозничнойЦены_БерётсяПредложениеУТЗ()
    {
        var grid = SheetGrid.FromRows(5, 2, new[]
        {
            new object?[] { "штрихкод", "Предложение УТЗ, руб." },
            new object?[] { 1d, 100d },
        });

        var table = ApprovedPriceSheetReader.FindTable(grid);

        Assert.Equal(3, table!.PriceColumn);
        Assert.Equal(5, table.HeaderRow);
    }

    [Fact]
    public void БезКолонкиЦены_ТаблицыНет()
    {
        var grid = SheetGrid.FromRows(1, 1, new[] { new object?[] { "ШК", "Артикул" } });

        Assert.Null(ApprovedPriceSheetReader.FindTable(grid));
    }
}

public class DuplicateBarcodeGroupsTests
{
    [Fact]
    public void ОдинаковыеШтрихкодыПолучаютОдинНомер_ОдиночкиБезНомера()
    {
        var groups = DuplicateBarcodeGroups.Assign(new[] { "111", "222", "333", "222", "111", "", "" });

        Assert.Equal(new int?[] { 0, 1, null, 1, 0, null, null }, groups);
    }

    [Fact]
    public void ЦветаРазныеИНеСовпадаютСПометками()
    {
        var palette = DuplicateBarcodeGroups.Palette;

        Assert.Equal(palette.Count, palette.Distinct().Count());
        Assert.DoesNotContain(0xCEC7FF, palette);
        Assert.DoesNotContain(0xCEEFC6, palette);
        Assert.Equal(DuplicateBarcodeGroups.ColorOf(0), DuplicateBarcodeGroups.ColorOf(palette.Count));
        Assert.NotEqual(DuplicateBarcodeGroups.ColorOf(0), DuplicateBarcodeGroups.ColorOf(1));
    }
}

public class SeasonalityMinimumsTests
{
    [Fact]
    public void ПрошлыеЗначенияУбираются_СовпавшиеПодгруппыСохраняются()
    {
        // Скрин: в новом списке одна подгруппа, ниже - «Мин раздача» прошлого распреда.
        var previous = new (string? Subgroup, object? Minimum)[]
        {
            ("МАСКА ДЛЯ СНА/", 2d),
            ("БРЕЛОК/", 3d),
            (null, 2d),
            ("", 1d),
        };

        var minimums = SeasonalityMinimums.Carry(previous, new[] { "брелок/", "ЗОНТ/" });

        Assert.Equal(new object?[] { 3d, null }, minimums);
    }
}
