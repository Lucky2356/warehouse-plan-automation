using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class NetworkDateParserTests
{
    private static readonly DateTime Today = new(2026, 9, 14);

    /// <summary>28.08.2026 - дата документа в числах Excel.</summary>
    private static readonly double DocumentDate = new DateTime(2026, 8, 28).ToOADate();

    [Theory]
    [InlineData("2452-010,014 Угги, кеды СЕТ 1 для Юга_в рознице с 15.10", "15.10", 2026, 10, 15)]
    [InlineData("2412-015 Пуховики_получение в рознице 20.09", "20.09", 2026, 9, 20)]
    [InlineData("2430-026, 028 Обувь СЕТ1_в рознице с 1.10", "1.10", 2026, 10, 1)]
    [InlineData("314-063 Сумки, рюкзаки_получение в розницет 01.08.2026", "01.08.2026", 2026, 8, 1)]
    [InlineData("Перчатки_ликвиды FW26-27_осн, доп мерч_в сети с 07.09_приемка на хранилище", "07.09", 2026, 9, 7)]
    public void ДатаИзТекстаПоставки(string text, string expectedText, int year, int month, int day)
    {
        var found = NetworkDateParser.Find(text, DocumentDate, Today);

        Assert.NotNull(found);
        Assert.Equal(expectedText, text.Substring(found!.Value.Start, found.Value.Length));
        Assert.Equal(new DateTime(year, month, day), found.Value.Date);
    }

    [Fact]
    public void ДатаДокумента_НеДатаВСети()
    {
        // «Срочная подтоварка 28.08» - день, когда завели заказ. В плане у таких строк
        // «Дата в сети» считается формулой от даты документа.
        Assert.Null(NetworkDateParser.Find("Срочная подтоварка 28.08_Хранение, хранилище", DocumentDate, Today));
    }

    [Theory]
    [InlineData("2446-001, 326-060 ШПП_отгрузка по готовности")]
    [InlineData("2430-026,028 Обувь_в рознице")]
    [InlineData("Сумка 318-435ЧЕРНЫЙ28.5x39см")]
    [InlineData("Весы 2.5 кг")]
    [InlineData("")]
    public void БезДаты(string text) =>
        Assert.Null(NetworkDateParser.Find(text, DocumentDate, Today));

    [Fact]
    public void ГодБлижайшийКДатеДокумента()
    {
        var december = new DateTime(2026, 12, 20).ToOADate();

        var found = NetworkDateParser.Find("Пуховики_в рознице с 15.01", december, Today);

        Assert.Equal(new DateTime(2027, 1, 15), found!.Value.Date);
    }
}
