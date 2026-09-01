using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Text;
using WarehousePlanAutomation.Tests.TestData;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class PlanUpdateBuilderTests
{
    private static readonly DateTime Monday = new(2026, 8, 31);
    private static readonly DateTime Saturday = new(2026, 9, 5);

    private static OrderRow Order(int excelRow, string division, string comment, double difference) =>
        new(excelRow, division, comment, difference, 46262.7d);

    private static JournalRow Journal(int order, long loadNumber, string status, double? percent) =>
        new(order, order + 2, "Заказ Номер загрузки " + loadNumber, status, percent);

    private static PlanStructuralUpdate Build(
        IReadOnlyList<OrderRow> orders,
        IReadOnlyList<JournalRow> journal,
        DateTime? today = null)
    {
        var layout = PlanFixture.BuildLayout();
        var classification = OrderClassifier.Classify(orders);
        return PlanUpdateBuilder.Build(layout, classification, journal, today ?? Monday);
    }

    [Fact]
    public void СуществующийЗаказ_ОбновляетсяБезДобавленияСтроки()
    {
        var update = Build(
            new[]
            {
                Order(2, "Москва-M149",
                    "Срочная подтоварка 28.08_Хранение Номер загрузки " + PlanFixture.ExistingUrgentLoadNumber, 900),
                Order(3, "Иркутск-М47",
                    "Срочная подтоварка 28.08_Хранение Номер загрузки " + PlanFixture.ExistingUrgentLoadNumber, 100),
            },
            new[] { Journal(0, PlanFixture.ExistingUrgentLoadNumber, "ЗАКРЫТ", 83) });

        Assert.Empty(update.NewRows);

        var order = update.OrderUpdates.Single(u => u.LoadNumber == PlanFixture.ExistingUrgentLoadNumber);
        Assert.Equal(1000d, order.Quantity);
        Assert.Equal(OrderTextRules.InAssemblyStatus, order.Status);
        Assert.Equal(83d, order.CompletionPercent);
    }

    [Fact]
    public void НовыйЗаказ_ДобавляетсяВБлокВсеГруппы()
    {
        const long newLoadNumber = 55600001;
        var update = Build(
            new[]
            {
                Order(2, "Москва-M149",
                    "1022-015 Пуховики_получение в рознице 20.09 Номер загрузки " + newLoadNumber + " <Подбор:>", 251),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН", 0) });

        var newRow = Assert.Single(update.NewRows);
        Assert.Equal(PlanSectionKind.AllGroups, newRow.Section);
        Assert.Equal("1022-015 Пуховики_получение в рознице 20.09", newRow.Supplies);
        Assert.Equal(OrderTextRules.LoadedComment, newRow.Comments);
        Assert.Equal(newLoadNumber, newRow.LoadNumber);
        Assert.Equal(46262d, newRow.DocumentDate);
    }

    [Fact]
    public void НовыйЗаказСНомеромПоставки_ПолучаетПерекр()
    {
        const long newLoadNumber = 55600002;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М139", "МЗ806-133 ШПП_отгрузка по готовности Номер загрузки " + newLoadNumber, 10),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН", 0) });

        Assert.Equal(OrderTextRules.CrossDockProcessing, update.NewRows.Single().Processing);
    }

    [Fact]
    public void НовыйЗаказБезНомераПоставки_ОстаётсяБезОбработки()
    {
        const long newLoadNumber = 55600003;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М139", "Бижутерия с хранилища_срочная отгрузка Номер загрузки " + newLoadNumber, 10),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН", 0) });

        Assert.Equal(string.Empty, update.NewRows.Single().Processing);
    }

    [Fact]
    public void НовыйЗаказИзВозвратов_ПопадаетВБлокВозвраты()
    {
        const long newLoadNumber = 55600004;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М139",
                    "Пуховики ликвиды_из возвратов, времянки Номер загрузки " + newLoadNumber, 86),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН", 0) });

        Assert.Equal(PlanSectionKind.Returns, update.NewRows.Single().Section);
    }

    [Fact]
    public void ИсчезнувшийЗаказ_ПолучаетНулевоеКоличество()
    {
        var update = Build(
            Array.Empty<OrderRow>(),
            new[] { Journal(0, PlanFixture.ExistingSetLoadNumber, "ЗАКРЫТ", 55) });

        var order = update.OrderUpdates.Single(u => u.LoadNumber == PlanFixture.ExistingSetLoadNumber);
        Assert.Equal(0d, order.Quantity);
    }

    [Fact]
    public void ЗаказаНетВЖурнале_КоличествоНольСтатусИПроцентНеМеняются()
    {
        var update = Build(
            new[]
            {
                Order(2, "Москва-M149",
                    "Подтоварка Номер загрузки " + PlanFixture.ExistingUrgentLoadNumber, 500),
            },
            Array.Empty<JournalRow>());

        var order = update.OrderUpdates.Single(u => u.LoadNumber == PlanFixture.ExistingUrgentLoadNumber);
        Assert.Equal(0d, order.Quantity);
        Assert.Null(order.Status);
        Assert.Null(order.CompletionPercent);
    }

    [Fact]
    public void СтрокаЗаказыБудутЗагружены_УдаляетсяПриПоявленииТойЖеПоставки()
    {
        const long newLoadNumber = 55600005;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М139", "1079-051 Шапки_отгрузка по готовности Номер загрузки " + newLoadNumber, 2000),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН", 0) });

        Assert.Equal(new[] { 8 }, update.PlanRowsToDelete);
    }

    [Fact]
    public void СтрокаЗаказыБудутЗагружены_ОстаётсяЕслиПоставкаНеПришла()
    {
        const long newLoadNumber = 55600006;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М139", "1244-001 ШПП_отгрузка по готовности Номер загрузки " + newLoadNumber, 100),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН", 0) });

        Assert.Empty(update.PlanRowsToDelete);
    }

    [Fact]
    public void ИтогиМаркетплейсовОптаИИнтернетМагазинаЗаписываются()
    {
        var update = Build(
            new[]
            {
                Order(2, "Опт", "Опт, обувь.", 4590),
                Order(3, "НОВОСИБИРСК-M-ИМ777", "Заказ интерент магазина № Т142347", 125),
                Order(4, "Lamoda", "Lamoda Подтоварка обувь зима из возвратов", 5281),
                Order(5, "Ozon", "Озон Екб МЗ1244-002 (Шапки) РФ", 13191),
                Order(6, "Wildberries", "WB Владимир ШПП приоритет к 24.07", 2339),
            },
            Array.Empty<JournalRow>());

        double Quantity(PlanAggregateTarget target) =>
            update.AggregateUpdates.Single(a => a.Target == target).Quantity;

        Assert.Equal(4590d, Quantity(PlanAggregateTarget.Wholesale));
        Assert.Equal(125d, Quantity(PlanAggregateTarget.InternetShop));
        Assert.Equal(5281d, Quantity(PlanAggregateTarget.MarketplaceFromReturns));
        Assert.Equal(13191d, Quantity(PlanAggregateTarget.MarketplaceFromSupplies));
        Assert.Equal(2339d, Quantity(PlanAggregateTarget.MarketplaceFromStorage));
    }

    [Fact]
    public void ПустыеКатегорииДаютНоль()
    {
        var update = Build(Array.Empty<OrderRow>(), Array.Empty<JournalRow>());

        Assert.All(
            update.AggregateUpdates.Where(a => a.Target != PlanAggregateTarget.AutoHub),
            aggregate => Assert.Equal(0d, aggregate.Quantity));
    }

    [Fact]
    public void АвтозаказыДляХабов_ВБудниЗаписываются()
    {
        var update = Build(Array.Empty<OrderRow>(), Array.Empty<JournalRow>(), Monday);

        Assert.Equal(24000d, update.AggregateUpdates.Single(a => a.Target == PlanAggregateTarget.AutoHub).Quantity);
    }

    [Fact]
    public void АвтозаказыДляХабов_ВВыходныеНеТрогаются()
    {
        var update = Build(Array.Empty<OrderRow>(), Array.Empty<JournalRow>(), Saturday);

        Assert.DoesNotContain(update.AggregateUpdates, a => a.Target == PlanAggregateTarget.AutoHub);
    }

    [Fact]
    public void ОсобыеСтрокиПланаНеПопадаютВОбновленияЗаказов()
    {
        var update = Build(Array.Empty<OrderRow>(), Array.Empty<JournalRow>());

        Assert.Equal(
            new[]
            {
                PlanFixture.ExistingUrgentLoadNumber,
                PlanFixture.ExistingSetLoadNumber,
                PlanFixture.ExistingMonoLoadNumber,
                PlanFixture.ExistingReturnLoadNumber,
                PlanFixture.ExistingStorageLoadNumber,
            },
            update.OrderUpdates.Select(u => u.LoadNumber));
    }
}
