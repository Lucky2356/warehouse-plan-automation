using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class StockSheetBuilderTests
{
    /// <summary>Выгрузка «Остатки Н» до ручной правки: с колонкой «АЦ» и лишними колонками.</summary>
    private static SheetGrid RawExport() => SheetGrid.FromRows(1, 1, new List<object?[]>
    {
        new object?[]
        {
            "АЦ", "Сектор", "Группа", "Артикул", "Размер", "Цвет", "Резерв",
            "Остаток 000 склад", "Остаток маркетплейс", "Без маркировки",
            "Остаток нет маркировки", "001 тз", "001 в_пути", "001 прод",
        },
        new object?[]
        {
            "314-047ЧЕРНЫЙ", "СУМКИ", "СУМКА", "314-047", "22х35см", "ЧЕРНЫЙ", 0d,
            11d, 0d, 0d, 0d, 1d, 2d, 3d,
        },
    });

    [Fact]
    public void ЛишниеКолонкиНеПереносятся()
    {
        StockSheetBuilder.ChooseColumns(RawExport(), 1, out var labels);

        Assert.DoesNotContain("АЦ", labels);
        Assert.DoesNotContain("Резерв", labels);
        Assert.DoesNotContain("Остаток маркетплейс", labels);
        Assert.DoesNotContain("Без маркировки", labels);
        Assert.DoesNotContain("Остаток нет маркировки", labels);
    }

    [Fact]
    public void АЦРСтоитПервойКолонкой()
    {
        StockSheetBuilder.ChooseColumns(RawExport(), 1, out var labels);

        Assert.Equal("АЦР", labels[0]);
        Assert.Equal(new[] { "АЦР", "Сектор", "Группа", "Артикул", "Размер", "Цвет", "Остаток 000 склад", "001 тз", "001 в_пути", "001 прод" }, labels);
    }

    [Fact]
    public void ГотовуюКолонкуАЦРПовторноНеДобавляет()
    {
        // Лист могли уже поправить руками - тогда «АЦР» в выгрузке есть.
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[] { "АЦР", "Сектор", "Артикул", "Размер", "Цвет" },
        });

        StockSheetBuilder.ChooseColumns(grid, 1, out var labels);

        Assert.Equal(new[] { "АЦР", "Сектор", "Артикул", "Размер", "Цвет" }, labels);
    }

    [Fact]
    public void АЦРСобираетсяИзАртикулаЦветаИРазмера()
    {
        Assert.Equal(
            "431-2802СЕРЫЙ-ЧЕРНЫЙTU",
            StockSheetBuilder.BuildAcr(" 431-2802 ", "СЕРЫЙ-ЧЕРНЫЙ", "TU"));
    }

    [Theory]
    [InlineData("001 тз", true)]
    [InlineData("001 в_пути", true)]
    [InlineData("001 прод", false)]
    [InlineData("Остаток 000 склад", false)]
    [InlineData("АЦР", false)]
    [InlineData("Наличие тз + в пути магазины", false)]
    public void ВОстатокИдутТолькоСтрокиТЗиВПути(string label, bool expected)
    {
        // Строка «прод» - это проданное, остатком оно не является: в разобранном файле
        // код магазина рядом с ней не проставлен, и «Распред» её не суммирует.
        var content = new StockSheetContent(new[] { label }, Array.Empty<StockSourceRow>());

        Assert.Equal(expected, content.IsStockRow(0));
    }
}
