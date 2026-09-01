using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class OrderClassifierTests
{
    private static OrderRow Row(int excelRow, string division, string comment, double difference) =>
        new(excelRow, division, comment, difference, 46262d);

    [Fact]
    public void Опт_СуммируетсяИСтрокиУдаляются()
    {
        var rows = new[]
        {
            Row(2, "Опт", "Опт, обувь, деми 2я часть.", 543),
            Row(3, "ОПТ", "Опт, сумки, рюкзаки.", 457),
            Row(4, "Москва-M149", "Подтоварка Номер загрузки 55575395", 10),
        };

        var result = OrderClassifier.Classify(rows);

        Assert.Equal(1000d, result.Total(OrderCategory.Wholesale));
        Assert.Contains(2, result.RowsToDelete);
        Assert.Contains(3, result.RowsToDelete);
    }

    [Fact]
    public void Опт_БезСтрокДаётНоль()
    {
        var result = OrderClassifier.Classify(new[] { Row(2, "Склад-А", "на образцы", 5) });

        Assert.Equal(0d, result.Total(OrderCategory.Wholesale));
    }

    [Fact]
    public void ИнтернетМагазин777_СуммируетсяИСтрокиУдаляются()
    {
        var rows = new[]
        {
            Row(2, "НОВОСИБИРСК-M-ИМ777", "Заказ интерент магазина № 0124589289-0331-1", 3),
            Row(3, "НОВОСИБИРСК-M-ИМ777", "Заказ интерент магазина № Т142347", 4),
        };

        var result = OrderClassifier.Classify(rows);

        Assert.Equal(7d, result.Total(OrderCategory.InternetShop));
        Assert.Equal(new[] { 2, 3 }, result.RowsToDelete);
        Assert.Empty(result.Leftovers);
    }

    [Theory]
    [InlineData("Ozon")]
    [InlineData("Озон")]
    [InlineData("Wildberries")]
    [InlineData("ВБ Владимир")]
    [InlineData("Lamoda")]
    [InlineData("Ламода")]
    [InlineData("Сбер Мега Маркет")]
    [InlineData("Екатеринбург Яблоко")]
    [InlineData("Магнит")]
    [InlineData("Магнит Москва")]
    [InlineData("ТС Магнит")]
    public void Маркетплейсы_РаспознаютсяВоВсехНаписаниях(string division)
    {
        var rows = new[] { Row(2, division, "Подтоварка микс приоритет к 10.08", 25) };

        var result = OrderClassifier.Classify(rows);

        Assert.Equal(25d, result.Total(OrderCategory.MarketplaceStorage));
        Assert.Empty(result.Leftovers);
    }

    [Theory]
    [InlineData("Магнитогорск-М96")]
    [InlineData("Магнитогорск-М12")]
    public void Магнитогорск_НеСчитаетсяМаркетплейсом(string division)
    {
        // «Магнит» ищется как отдельное слово, иначе обычный магазин Магнитогорска
        // попал бы в итоги маркетплейсов.
        var rows = new[]
        {
            Row(2, division, "Подтоварка ШПП Номер загрузки 55575395 <Подбор:>", 40),
        };

        var result = OrderClassifier.Classify(rows);

        Assert.Equal(0d, result.Total(OrderCategory.MarketplaceStorage));
        Assert.Equal(0d, result.Total(OrderCategory.MarketplaceSupplies));
        Assert.Equal(0d, result.Total(OrderCategory.MarketplaceReturns));
        var group = Assert.Single(result.Groups);
        Assert.Equal(55575395L, group.LoadNumber);
        Assert.Equal(40d, group.Quantity);
    }

    [Fact]
    public void Маркетплейсы_ВозвратыОбрабатываютсяРаньшеПоставок()
    {
        // В комментарии одновременно есть «возвр» и номер поставки: побеждает правило возвратов.
        var rows = new[]
        {
            Row(2, "Ozon", "Ozon Мск Крупный микс из Возвратов и С139-058 приоритет к 06.08", 40),
        };

        var result = OrderClassifier.Classify(rows);

        Assert.Equal(40d, result.Total(OrderCategory.MarketplaceReturns));
        Assert.Equal(0d, result.Total(OrderCategory.MarketplaceSupplies));
    }

    [Fact]
    public void Маркетплейсы_РазделяютсяНаВозвратыПоставкиИХранение()
    {
        var rows = new[]
        {
            Row(2, "Lamoda", "Lamoda Подтоварка обувь зима из возвратов", 10),
            Row(3, "Ozon", "Озон Екб МЗ1244-002, МЗ798-151 (Шапки) РФ", 20),
            Row(4, "Wildberries", "WB Владимир ШПП приоритет к 24.07", 30),
        };

        var result = OrderClassifier.Classify(rows);

        Assert.Equal(10d, result.Total(OrderCategory.MarketplaceReturns));
        Assert.Equal(20d, result.Total(OrderCategory.MarketplaceSupplies));
        Assert.Equal(30d, result.Total(OrderCategory.MarketplaceStorage));
        Assert.Equal(new[] { 2, 3, 4 }, result.RowsToDelete);
    }

    [Theory]
    [InlineData("Заказ на магазин 001 из заказов МП автозаказ Номер загрузки 55386696")]
    [InlineData("СЗ798-154(шапки), почта Росс, виртуально, 1 место Номер загрузки 55050468")]
    [InlineData("на фото, оставят в офисе для работы, со склада списать арт.140-349")]
    [InlineData("на образцы, хранение на складе СЗ985-062 (тапки)")]
    [InlineData("1180-008 в ОФис на ремонт с возвратом.")]
    [InlineData("В Офис 1180-008 Ремонт с возвратом WB")]
    [InlineData("Укомплектовать в наборы МЗ541-027 (палантин+брошь). Списать со склада.")]
    [InlineData("Маркировка.")]
    [InlineData("Виртуальный возврат")]
    public void СлужебныеСтроки_УдаляютсяБезПереноса(string comment)
    {
        var rows = new[] { Row(2, "Новосибирск-M77a", comment, 12) };

        var result = OrderClassifier.Classify(rows);

        Assert.Equal(12d, result.Total(OrderCategory.Service));
        Assert.Contains(2, result.RowsToDelete);
        Assert.Empty(result.Groups);
    }

    [Fact]
    public void РеальныеЗаказы_ГруппируютсяПоНомеруЗагрузки()
    {
        var rows = new[]
        {
            Row(2, "Москва-M149", "Срочная подтоварка 28.08_Хранение Номер загрузки 55575395 <Подбор:>", 100),
            Row(3, "Иркутск-М47", "Срочная подтоварка 28.08_Хранение Номер загрузки 55575395 <Подбор:>", 50),
            Row(4, "Казань-М139", "1022-015 Пуховики Номер загрузки 55231895 <Подбор:>", 7),
        };

        var result = OrderClassifier.Classify(rows);

        Assert.Equal(2, result.Groups.Count);
        var first = result.Groups[0];
        Assert.Equal(55575395L, first.LoadNumber);
        Assert.Equal(150d, first.Quantity);
        Assert.Equal(new[] { 2, 3 }, first.SourceRows);
        Assert.Equal(new[] { 2, 3, 4 }, result.RowsToDelete);
    }

    [Fact]
    public void СтрокиБезНомераЗагрузки_ОстаютсяДляРучнойПроверки()
    {
        // Ни служебного слова, ни номера загрузки: строку разбирают вручную.
        var rows = new[]
        {
            Row(2, "Склад-А", "Перемещение между зонами хранения", 5),
            Row(3, "Москва-M149", "Подтоварка Номер загрузки 55575395", 8),
        };

        var result = OrderClassifier.Classify(rows);

        Assert.Single(result.Leftovers);
        Assert.Equal(2, result.Leftovers[0].ExcelRow);
        Assert.DoesNotContain(2, result.RowsToDelete);
        Assert.Contains(3, result.RowsToDelete);
    }
}
