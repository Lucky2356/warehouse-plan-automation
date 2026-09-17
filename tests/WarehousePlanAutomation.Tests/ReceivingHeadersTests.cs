using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

/// <summary>
/// Листы «А2, А3», «МП» и листы поставок программа заполняет сама, включая названия колонок:
/// разбор не должен останавливаться из-за того, что колонок данных в книге ещё нет.
/// </summary>
public class ReceivingHeadersTests
{
    private static SheetGrid Row(params object?[] headers) =>
        SheetGrid.FromRows(1, 1, new List<object?[]> { headers });

    [Fact]
    public void Хранение_ТолькоКолонкиФормул_КнигаРазбирается()
    {
        var headers = HeaderResolver.Resolve(
            Row("В приемку", "Запас нарастающим", "Учитывать", "Номер адреса", "Номер тары"),
            ReceivingSchema.StorageSheet,
            ReceivingSchema.Storage.Specs);

        Assert.Equal(1, headers[ReceivingSchema.Storage.ToReceive]);
        Assert.Equal(3, headers[ReceivingSchema.Storage.Counted]);
        Assert.False(headers.TryGet(ReceivingSchema.Storage.Address, out _));
        Assert.False(headers.TryGet(ReceivingSchema.Storage.Code, out _));
    }

    [Fact]
    public void Хранение_БезКолонкиФормулы_Останавливается()
    {
        // «Учитывать» и номера программа не сочиняет: без них разбирать нечего.
        var error = Assert.Throws<WorkbookValidationException>(() => HeaderResolver.Resolve(
            Row("Адрес", "Тара", "Артикул", "Код", "В приемку", "Номер адреса", "Номер тары"),
            ReceivingSchema.StorageSheet,
            ReceivingSchema.Storage.Specs));

        Assert.Contains(ReceivingSchema.Storage.Counted, Assert.Single(error.Problems));
    }

    [Fact]
    public void Хранение_ГотоваяКнига_ВсеКолонкиНаходятся()
    {
        var headers = HeaderResolver.Resolve(
            Row(
                "Адрес", "Тара", "Артикул", "Код", "Сектор", "Группа", "Подгруппа", "наименование",
                "Качество", "Количество", "Размер", "Сезон", "Тема", "Цвет",
                "В приемку", "Запас нарастающим", "Учитывать", "Номер адреса", "Номер тары"),
            ReceivingSchema.StorageSheet,
            ReceivingSchema.Storage.Specs);

        Assert.Equal(1, headers[ReceivingSchema.Storage.Address]);
        Assert.Equal(14, headers[ReceivingSchema.Storage.Color]);
        Assert.Equal(15, headers[ReceivingSchema.Storage.ToReceive]);
        Assert.Equal(19, headers[ReceivingSchema.Storage.ContainerNumber]);
    }

    [Fact]
    public void Хранение_ПорядокСозданияПовторяетЛист()
    {
        var titles = ReceivingSchema.Storage.Created.Select(column => column.Title).ToList();

        Assert.Equal(ReceivingSchema.Storage.Address, titles[0]);
        Assert.Equal(ReceivingSchema.Storage.Color, titles[13]);
        Assert.Equal(ReceivingSchema.Storage.ToReceive, titles[14]);
        Assert.Equal(ReceivingSchema.Storage.ContainerNumber, titles[^1]);
        Assert.Equal(titles.Count, titles.Distinct().Count());
    }

    [Fact]
    public void Поставки_ТолькоКолонкиФормул_КнигаРазбирается()
    {
        var headers = HeaderResolver.Resolve(
            Row("Код", "Тара/код поставщика"),
            ReceivingSchema.CollectedSheet,
            ReceivingSchema.Supplies.Specs);

        Assert.Equal(1, headers[ReceivingSchema.Supplies.Code]);
        Assert.Equal(2, headers[ReceivingSchema.Supplies.ContainerCode]);
        Assert.False(headers.TryGet(ReceivingSchema.Supplies.SupplyNumber, out _));
        Assert.False(headers.TryGet(ReceivingSchema.Supplies.Barcode, out _));
    }

    [Fact]
    public void Поставки_БезКолонкиФормулы_Останавливается()
    {
        var error = Assert.Throws<WorkbookValidationException>(() => HeaderResolver.Resolve(
            Row("Номер поставки", "Шк"),
            ReceivingSchema.CollectedSheet,
            ReceivingSchema.Supplies.Specs));

        Assert.Equal(2, error.Problems.Count);
    }
}
