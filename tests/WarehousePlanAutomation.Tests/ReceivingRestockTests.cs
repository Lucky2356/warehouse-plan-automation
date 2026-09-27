using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

/// <summary>Приемка на хранилище: «Допоставить», которое заполняет первый этап.</summary>
public class ReceivingRestockTests
{
    private static RestockState Row(
        object? restock = null,
        double remainder = 0,
        double sold = 0,
        double storage = 0,
        double collected = 0,
        double notCollected = 0,
        double marketplace = 0,
        params object?[] stenki) =>
        new(0, restock, remainder, sold, storage, collected, notCollected, marketplace,
            stenki.Length == 0 ? new object?[] { "07/09-31/01", null, null } : stenki);

    private static RestockChoice? Plan(RestockState row) => ReceivingSummary.PlanRestock(new[] { row }).SingleOrDefault();

    [Fact]
    public void НетОстатковПродажИСобранных_Ноль()
    {
        // Даже если что-то лежит в несобранных поставках: по инструкции товар без продаж
        // и остатков не берём.
        var choice = Plan(Row(notCollected: 30));

        Assert.Equal((RestockSource.Zero, 0d), (choice!.Source, choice.Value));
    }

    [Fact]
    public void Ноль_НеСтавитсяТам_ГдеЕстьКоличествоМП()
    {
        // Пустое «Допоставить» таких строк на втором этапе получает «Количество МП».
        Assert.Null(Plan(Row(notCollected: 30, marketplace: 12, stenki: new object?[] { "МП", null, null })));
    }

    [Fact]
    public void ЕстьНаХранении_БерёмВсёСА2А3()
    {
        var choice = Plan(Row(sold: 5, storage: 17, collected: 40));

        Assert.Equal((RestockSource.Storage, 17d), (choice!.Source, choice.Value));
    }

    [Fact]
    public void НетНаХранении_БерёмВсёСобранное()
    {
        var choice = Plan(Row(sold: 6, collected: 184, notCollected: 20));

        Assert.Equal((RestockSource.Collected, 184d), (choice!.Source, choice.Value));
    }

    [Fact]
    public void ТолькоНесобранные_БерёмИх()
    {
        var choice = Plan(Row(sold: 3, notCollected: 25));

        Assert.Equal((RestockSource.NotCollected, 25d), (choice!.Source, choice.Value));
    }

    [Fact]
    public void НаХраненииДесятьИМеньше_БерёмСобранное()
    {
        var choice = Plan(Row(sold: 6, storage: 10, collected: 40));

        Assert.Equal((RestockSource.Collected, 40d), (choice!.Source, choice.Value));
    }

    [Fact]
    public void СобранныхДесятьИМеньше_БерёмНесобранные()
    {
        var choice = Plan(Row(sold: 6, storage: 4, collected: 9, notCollected: 11));

        Assert.Equal((RestockSource.NotCollected, 11d), (choice!.Source, choice.Value));
    }

    [Fact]
    public void ВездеДесятьИМеньше_ОставляемПустым()
    {
        Assert.Null(Plan(Row(sold: 6, storage: 10, collected: 7, notCollected: 3)));
    }

    [Theory]
    [InlineData("МП", null, null)]
    [InlineData(null, "МП", null)]
    [InlineData(null, null, "мп 01.10")]
    public void МПВСтенках_НеТрогаем(string? main, string? small, string? extra)
    {
        Assert.Null(Plan(Row(sold: 6, storage: 17, stenki: new object?[] { main, small, extra })));
    }

    [Fact]
    public void ОшибкаИДатыВСтенках_НеМешают()
    {
        var choice = Plan(Row(sold: 6, collected: 50, stenki: new object?[] { -2146826246, "01.11.2026-31.01.2027", "компл" }));

        Assert.Equal(RestockSource.Collected, choice!.Source);
    }

    [Theory]
    [InlineData(40d)]
    [InlineData(0d)]
    [InlineData("решу завтра")]
    public void Заполненное_НеТрогаем(object restock)
    {
        Assert.Null(Plan(Row(restock: restock, sold: 6, collected: 184)));
    }
}
