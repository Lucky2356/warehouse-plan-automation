using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class JournalEvaluatorTests
{
    private const long LoadNumber = 55575395;

    private static JournalRow Row(int order, string status, double? percent) =>
        new(order, order + 2, "Срочная подтоварка 28.08_Хранение Номер загрузки " + LoadNumber, status, percent);

    [Fact]
    public void НетЗаписиВЖурнале_ЗаказНеНайден()
    {
        var journal = new[]
        {
            new JournalRow(0, 2, "Другой заказ Номер загрузки 55377185", "ЗАКРЫТ", 100),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.False(outcome.Found);
        Assert.False(outcome.SetInAssembly);
        Assert.Null(outcome.PercentToSet);
    }

    [Fact]
    public void ПервоеВхождение_ЭтоСамаяВерхняяСтрокаЖурнала()
    {
        // Журнал не сортируется: значение берётся из первой физической строки, а не из максимума.
        var journal = new[]
        {
            Row(0, "ЗАКРЫТ", 83),
            Row(1, "ЗАКРЫТ", 100),
            Row(2, "ЗАПУЩЕН", 0),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.Found);
        Assert.True(outcome.SetInAssembly);
        Assert.Equal(83d, outcome.PercentToSet);
    }

    [Fact]
    public void ПроцентБольшеНуля_СтатусВСборкеИПроцентИзЖурнала()
    {
        var journal = new[] { Row(0, "ЗАКРЫТ", 68) };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.SetInAssembly);
        Assert.Equal(68d, outcome.PercentToSet);
    }

    [Fact]
    public void ПроцентСто_СтатусВсёРавноВСборке()
    {
        var journal = new[] { Row(0, "ЗАКРЫТ", 100) };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.SetInAssembly);
        Assert.Equal(100d, outcome.PercentToSet);
    }

    [Fact]
    public void ПервыйПроцентНоль_ЕстьЗапущен_СтатусВСборкеБезИзмененияПроцента()
    {
        var journal = new[]
        {
            Row(0, "ЗАКРЫТ", 0),
            Row(1, "ПРОВЕДЕН", 0),
            Row(2, "Запущен", 0),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.Found);
        Assert.True(outcome.SetInAssembly);
        Assert.Null(outcome.PercentToSet);
    }

    [Fact]
    public void ПервыйПроцентНоль_НетЗапущен_СтатусИПроцентНеМеняются()
    {
        var journal = new[]
        {
            Row(0, "ЗАКРЫТ", 0),
            Row(1, "ПРОВЕДЕН", 0),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.Found);
        Assert.False(outcome.SetInAssembly);
        Assert.Null(outcome.PercentToSet);
    }

    [Fact]
    public void ПустойПроцентСчитаетсяНулём()
    {
        var journal = new[]
        {
            Row(0, "ЗАКРЫТ", null),
            Row(1, "ЗАПУЩЕН", null),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.SetInAssembly);
        Assert.Null(outcome.PercentToSet);
    }
}
