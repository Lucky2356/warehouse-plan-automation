using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class JournalEvaluatorTests
{
    private const long LoadNumber = 55388195;

    /// <summary>Строка документа магазина: она и участвует в подсчёте выполнения.</summary>
    private static JournalRow Doc(
        int order,
        string number,
        double planned,
        double actual,
        string status = "ЗАПУЩЕН") =>
        new(
            order,
            order + 2,
            "806-132, 799-041 ШПП_отгрузка по готовности Номер загрузки " + LoadNumber,
            status,
            planned > 0 ? Math.Round(actual / planned * 100) : 0,
            number,
            planned,
            actual);

    /// <summary>
    /// Сводная строка группы: в колонке «Номер» стоит число, а «Кол-во ед» уже включает
    /// строки магазинов. В подсчёт такая строка попадать не должна.
    /// </summary>
    private static JournalRow Summary(int order, double planned, double actual) =>
        new(
            order,
            order + 2,
            "806-132, 799-041 ШПП_отгрузка по готовности Номер загрузки " + LoadNumber,
            "ЗАКРЫТ",
            100,
            "-21",
            planned,
            actual);

    [Fact]
    public void НетЗаписиВЖурнале_ЗаказНеНайден()
    {
        var journal = new[]
        {
            new JournalRow(0, 2, "Другой заказ Номер загрузки 55377185", "ЗАКРЫТ", 100, "З000-1", 10, 10),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.False(outcome.Found);
        Assert.False(outcome.SetInAssembly);
        Assert.Null(outcome.PercentToSet);
    }

    [Fact]
    public void ПроцентСчитаетсяПоСуммамКоличестваИФакта()
    {
        // Числа из реального журнала за 02.09: закрытая группа 8046 из 8046
        // и запущенная 2308 из 15427. Итого 10354 из 23473 - это 44 %,
        // хотя в колонке «%» самого журнала числа 44 нет ни в одной строке.
        var journal = new[]
        {
            Doc(0, "З000-260355", 8046, 8046, "ЗАКРЫТ"),
            Doc(1, "З000-260437", 15427, 2308),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.Found);
        Assert.Equal(44d, outcome.PercentToSet);
        Assert.True(outcome.SetInAssembly);
    }

    [Fact]
    public void СводныеСтрокиВРасчётНеПопадают()
    {
        // Сводная строка повторяет итог группы. Если её посчитать, выполнение удвоится.
        var journal = new[]
        {
            Summary(0, 8046, 8046),
            Doc(1, "З000-260355", 8046, 8046, "ЗАКРЫТ"),
            Doc(2, "З000-260437", 15427, 2308),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.Equal(44d, outcome.PercentToSet);
    }

    [Fact]
    public void ПовторыОдногоДокументаСчитаютсяОдинРаз()
    {
        // Лист журнала содержит несколько подряд приклеенных выгрузок, и закрытые
        // документы повторяются в каждой. Без дедупликации старые группы перевесили бы.
        var journal = new[]
        {
            Doc(0, "З000-260355", 8046, 8046, "ЗАКРЫТ"),
            Doc(1, "З000-260355", 8046, 8046, "ЗАКРЫТ"),
            Doc(2, "З000-260355", 8046, 8046, "ЗАКРЫТ"),
            Doc(3, "З000-260437", 15427, 2308),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.Equal(44d, outcome.PercentToSet);
    }

    [Fact]
    public void ПроцентОкругляетсяДоЦелого()
    {
        // 240 из 271 - это 88,56 %, в плане стоит 89.
        var journal = new[] { Doc(0, "З000-1", 271, 240) };

        Assert.Equal(89d, JournalEvaluator.Evaluate(LoadNumber, journal).PercentToSet);
    }

    [Fact]
    public void ПочтиНулевоеВыполнениеДаётНоль()
    {
        // 82 из 18847 - это 0,44 %, в плане стоит 0.
        var journal = new[] { Doc(0, "З000-1", 18847, 82) };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.Equal(0d, outcome.PercentToSet);
        Assert.True(outcome.SetInAssembly, "статус «Запущен» ставит заказ в сборку и при нулевом проценте");
    }

    [Fact]
    public void ПолноеВыполнение_СтатусВсёРавноВСборке()
    {
        var journal = new[] { Doc(0, "З000-1", 500, 500, "ЗАКРЫТ") };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.Equal(100d, outcome.PercentToSet);
        Assert.True(outcome.SetInAssembly);
    }

    [Fact]
    public void НулевоеВыполнениеБезЗапущен_СтатусНеМеняется()
    {
        var journal = new[]
        {
            Doc(0, "З000-1", 3437, 0, "ПРОВЕДЕН"),
            Doc(1, "З000-2", 1000, 0, "ПРОВЕДЕН"),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.Found);
        Assert.False(outcome.SetInAssembly);
        Assert.Equal(0d, outcome.PercentToSet);
    }

    [Fact]
    public void ЗаказБезСтрокДокументов_ДаётНоль()
    {
        var journal = new[] { Summary(0, 8046, 8046) };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.Found);
        Assert.Equal(0d, outcome.PercentToSet);
    }

    [Fact]
    public void НомерЗагрузкиНеНаходитсяВнутриБолееДлинногоЧисла()
    {
        // «55388195» не должно совпадать с «155388195» и «553881950»:
        // иначе заказу достался бы процент чужого заказа.
        var journal = new[]
        {
            new JournalRow(0, 2, "Подтоварка Номер загрузки 155388195", "ЗАКРЫТ", 90, "З000-1", 10, 9),
            new JournalRow(1, 3, "Подтоварка Номер загрузки 553881950", "ЗАПУЩЕН", 0, "З000-2", 10, 0),
        };

        Assert.False(JournalEvaluator.Evaluate(LoadNumber, journal).Found);
    }

    [Fact]
    public void ПустыеКоличестваСчитаютсяНулём()
    {
        var journal = new[]
        {
            new JournalRow(0, 2, "Заказ Номер загрузки " + LoadNumber, "ЗАПУЩЕН", null, "З000-1", null, null),
        };

        var outcome = JournalEvaluator.Evaluate(LoadNumber, journal);

        Assert.True(outcome.Found);
        Assert.Equal(0d, outcome.PercentToSet);
        Assert.True(outcome.SetInAssembly);
    }
}
