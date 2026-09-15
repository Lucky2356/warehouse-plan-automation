using System.Globalization;
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

    /// <summary>
    /// Строка журнала. Процент считается по «Кол-во ед» и «Факт ед», поэтому именно они
    /// задаются здесь, а не готовое значение колонки «%».
    /// </summary>
    private static JournalRow Journal(
        int order,
        long loadNumber,
        string status,
        double planned = 100d,
        double actual = 0d) =>
        new(
            order,
            order + 2,
            "Заказ Номер загрузки " + loadNumber,
            status,
            null,
            "З000-" + order.ToString(CultureInfo.InvariantCulture),
            planned,
            actual);

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
                Order(2, "Москва-M614",
                    "Срочная подтоварка 28.08_Хранение Номер загрузки " + PlanFixture.ExistingUrgentLoadNumber, 900),
                Order(3, "Иркутск-М57",
                    "Срочная подтоварка 28.08_Хранение Номер загрузки " + PlanFixture.ExistingUrgentLoadNumber, 100),
            },
            new[] { Journal(0, PlanFixture.ExistingUrgentLoadNumber, "ЗАКРЫТ", 100d, 83d) });

        Assert.Empty(update.NewRows);

        var order = update.OrderUpdates.Single(u => u.LoadNumber == PlanFixture.ExistingUrgentLoadNumber);
        Assert.Equal(1000d, order.Quantity);
        Assert.Equal(OrderTextRules.InAssemblyStatus, order.Status);
        Assert.Equal(83d, order.CompletionPercent);
    }

    [Fact]
    public void НовыйЗаказ_ДобавляетсяВБлокВсеГруппы()
    {
        const long newLoadNumber = 44600001;
        var update = Build(
            new[]
            {
                Order(2, "Москва-M614",
                    "2412-015 Пуховики_получение в рознице 20.09 Номер загрузки " + newLoadNumber + " <Подбор:>", 251),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        var newRow = Assert.Single(update.NewRows);
        Assert.Equal(PlanSectionKind.AllGroups, newRow.Section);
        Assert.Equal("2412-015 Пуховики_получение в рознице 20.09", newRow.Supplies);
        Assert.Equal(OrderTextRules.LoadedComment, newRow.Comments);
        Assert.Equal(newLoadNumber, newRow.LoadNumber);
        Assert.Equal(46262d, newRow.DocumentDate);
    }

    [Fact]
    public void НовыйЗаказСНомеромПоставки_ПолучаетПерекр()
    {
        const long newLoadNumber = 44600002;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613", "МЗ352-133 ШПП_отгрузка по готовности Номер загрузки " + newLoadNumber, 10),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        Assert.Equal(OrderTextRules.CrossDockProcessing, update.NewRows.Single().Processing);
    }

    [Fact]
    public void НовыйЗаказБезНомераПоставки_ОстаётсяБезОбработки()
    {
        const long newLoadNumber = 44600003;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613", "Бижутерия с хранилища_срочная отгрузка Номер загрузки " + newLoadNumber, 10),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        Assert.Equal(string.Empty, update.NewRows.Single().Processing);
    }

    [Fact]
    public void НовыйЗаказ_ПерекрОпределяетсяТолькоПоТекстуПоставок()
    {
        // Номер поставки стоит после слов «Номер загрузки», то есть в «Поставки» не попадает.
        // Признак «перекр» относится к поставке, а не к служебному хвосту комментария.
        const long newLoadNumber = 44600004;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613",
                    "Бижутерия с хранилища Номер загрузки " + newLoadNumber + " <Подбор: 2412-015>", 10),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        var newRow = update.NewRows.Single();
        Assert.Equal("Бижутерия с хранилища", newRow.Supplies);
        Assert.Equal(string.Empty, newRow.Processing);
    }

    [Fact]
    public void НовыйЗаказИзВозвратов_ПопадаетВБлокВозвраты()
    {
        const long newLoadNumber = 44600004;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613",
                    "Пуховики ликвиды_из возвратов, времянки Номер загрузки " + newLoadNumber, 86),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        Assert.Equal(PlanSectionKind.Returns, update.NewRows.Single().Section);
    }

    [Fact]
    public void НовыйЗаказСПриемкой_ПопадаетВБлокПриемкаНаХранилище()
    {
        const long newLoadNumber = 44600011;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613",
                    "Сезонный товар FW26-27_приёмка на хранилище_1 приоритет Номер загрузки " + newLoadNumber, 661),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        Assert.Equal(PlanSectionKind.StorageAcceptance, update.NewRows.Single().Section);
    }

    [Fact]
    public void ПриемкаИзВозвратов_ПриемкаСильнееВозврата()
    {
        // В плане такой заказ стоит в блоке «приемка на хранилище», хотя «возвратов» в нём тоже есть.
        const long newLoadNumber = 44600012;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613",
                    "Зимняя обувь, сумки_ликвиды FW26-27_получение в рознице 01.10_приемка на хранилище из возвратов " +
                    "Номер загрузки " + newLoadNumber, 670),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        var newRow = update.NewRows.Single();
        Assert.Equal(PlanSectionKind.StorageAcceptance, newRow.Section);
        Assert.Equal(new DateTime(2026, 10, 1), newRow.NetworkDate!.Value.Date);
    }

    [Fact]
    public void ДатаВТекстеПоставки_СтановитсяДатойВСети()
    {
        const long newLoadNumber = 44600013;
        var update = Build(
            new[]
            {
                Order(2, "Краснодар-М71",
                    "2452-010,014 Угги, кеды СЕТ 1 для Юга_в рознице с 15.10 Номер загрузки " + newLoadNumber, 84),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        var newRow = update.NewRows.Single();
        var date = newRow.NetworkDate!.Value;
        Assert.Equal(new DateTime(2026, 10, 15), date.Date);
        Assert.Equal("15.10", newRow.Supplies.Substring(date.Start, date.Length));
    }

    [Fact]
    public void НовыеСтроки_СрочныеПоДатеСетМоноОстальное()
    {
        // Скрин: «…СЕТ2_в рознице с 08.10» пришла раньше «…СЕТ2_в рознице с 01.10» и встала выше.
        // Правило аналитика (02.09): срочные выше, дальше по дате в сети, в одной дате -
        // сначала СЕТ, потом МОНО, потом остальное. Строки без даты - первыми.
        var comments = new[]
        {
            "2430-030, 033 Угги СЕТ2_в рознице с 08.10",
            "2430-030, 033 Угги МОНО_в рознице с 01.10",
            "2430-030, 033 Угги CET2_в рознице с 01.10",
            "377-001 Перчатки_отгрузка по готовности",
            "2430-030, 033 Угги СЕТ 1_в рознице с 01.10",
            "2430-030, 033 Угги_в рознице с 01.10",
            "Бижутерия с хранилища_срочная отгрузка",
        };

        var orders = comments
            .Select((comment, i) => Order(2 + i, "Москва-M614", comment + " Номер загрузки " + (44600100 + i), 10))
            .ToList();
        var journal = comments.Select((_, i) => Journal(i, 44600100 + i, "ЗАПУЩЕН")).ToList();

        var update = Build(orders, journal);

        Assert.Equal(
            new[]
            {
                "Бижутерия с хранилища_срочная отгрузка",
                "377-001 Перчатки_отгрузка по готовности",
                "2430-030, 033 Угги СЕТ 1_в рознице с 01.10",
                "2430-030, 033 Угги CET2_в рознице с 01.10",
                "2430-030, 033 Угги МОНО_в рознице с 01.10",
                "2430-030, 033 Угги_в рознице с 01.10",
                "2430-030, 033 Угги СЕТ2_в рознице с 08.10",
            },
            update.NewRows.Select(row => row.Supplies));
    }

    [Theory]
    [InlineData("Обувь СЕТ1 для Юга_в рознице с 15.10", 0)]
    [InlineData("Угги CET2_в рознице с 08.10", 0)]
    [InlineData("Обувь МОНО для Юга_в рознице с 15.10", 1)]
    [InlineData("Кеды MOHO_в рознице с 15.09", 1)]
    [InlineData("Обувь_в рознице с 01.10", 2)]
    [InlineData("Перчатки_в сети с 07.09", 2)]
    [InlineData("Кассета для сетевых магазинов", 2)]
    public void СетМоноОстальное(string supplies, int rank)
    {
        Assert.Equal(rank, NewPlanRowOrder.SetMonoRank(supplies));
    }

    [Fact]
    public void БезДатыВТексте_ДатуВСетиСчитаетФормула()
    {
        const long newLoadNumber = 44600014;
        var update = Build(
            new[]
            {
                Order(2, "Москва-M614", "Срочная подтоварка 28.08_Хранение Номер загрузки " + newLoadNumber, 900),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        Assert.Null(update.NewRows.Single().NetworkDate);
    }

    [Fact]
    public void НоваяСтрока_ПомечаетсяКакДобавленнаяСегодня()
    {
        const long newLoadNumber = 44600009;
        var update = Build(
            new[]
            {
                Order(2, "Москва-M614",
                    "2412-015 Пуховики_в рознице 20.09 Номер загрузки " + newLoadNumber, 251),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        var added = update.OrderUpdates.Single(u => u.LoadNumber == newLoadNumber);
        Assert.True(added.IsNewRow);
        Assert.False(added.MissingFromOrders);

        // Вчерашние строки пометку не получают, иначе зелёным окрасился бы весь план.
        Assert.All(
            update.OrderUpdates.Where(u => u.LoadNumber != newLoadNumber),
            u => Assert.False(u.IsNewRow));
    }

    [Fact]
    public void ЗапланированнаяСтрока_НеСчитаетсяНовой()
    {
        // Строку завела аналитик, программа только вписала в неё номер загрузки:
        // помечать её как добавленную сегодня нечестно.
        const long newLoadNumber = 44600010;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613",
                    PlanFixture.PlaceholderSupplies + " Номер загрузки " + newLoadNumber, 2000),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        Assert.Single(update.PlannedMatches);
        Assert.False(update.OrderUpdates.Single(u => u.LoadNumber == newLoadNumber).IsNewRow);
    }

    [Fact]
    public void ИсчезнувшийЗаказ_ПолучаетНулевоеКоличество()
    {
        var update = Build(
            Array.Empty<OrderRow>(),
            new[] { Journal(0, PlanFixture.ExistingSetLoadNumber, "ЗАКРЫТ", 100d, 55d) });

        var order = update.OrderUpdates.Single(u => u.LoadNumber == PlanFixture.ExistingSetLoadNumber);
        Assert.Equal(0d, order.Quantity);
    }

    [Fact]
    public void ЗаказаНетВЖурнале_КоличествоНольСтатусИПроцентНеМеняются()
    {
        var update = Build(
            new[]
            {
                Order(2, "Москва-M614",
                    "Подтоварка Номер загрузки " + PlanFixture.ExistingUrgentLoadNumber, 500),
            },
            Array.Empty<JournalRow>());

        var order = update.OrderUpdates.Single(u => u.LoadNumber == PlanFixture.ExistingUrgentLoadNumber);
        Assert.Equal(0d, order.Quantity);
        Assert.Null(order.Status);
        Assert.Null(order.CompletionPercent);
    }

    [Fact]
    public void СтрокаЗаказыБудутЗагружены_УдаляетсяПриПоявленииТойЖеПоставкиПодДругимНазванием()
    {
        // Заказ пришёл под своим названием, а заготовка называлась иначе: та же поставка
        // теперь представлена новой строкой, и заготовка не нужна.
        const long newLoadNumber = 44600005;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613", "2421-051 Шапки СЕТ1_в рознице с 10.09 Номер загрузки " + newLoadNumber, 2000),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        Assert.Equal(new[] { 8 }, update.PlanRowsToDelete);
        Assert.Single(update.NewRows);
        Assert.Empty(update.PlannedMatches);
    }

    [Fact]
    public void ЗапланированнаяСтрока_ПолучаетНомерЗагрузкиВместоВторойСтроки()
    {
        // Аналитик заводит будущую поставку строкой без номера загрузки. Когда заказ
        // приходит с тем же текстом, номер вписывается в неё, а не создаётся дубль:
        // иначе итог блока завышается ровно на количество этого заказа.
        const long newLoadNumber = 44600007;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613",
                    PlanFixture.PlaceholderSupplies + " Номер загрузки " + newLoadNumber, 2000),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        var match = Assert.Single(update.PlannedMatches);
        Assert.Equal(8, match.ExcelRow);
        Assert.Equal(newLoadNumber, match.LoadNumber);

        Assert.Empty(update.NewRows);
        Assert.Empty(update.PlanRowsToDelete);

        var order = update.OrderUpdates.Single(u => u.LoadNumber == newLoadNumber);
        Assert.Equal(2000d, order.Quantity);
    }

    [Fact]
    public void ЗапланированнаяСтрока_НеПодставляетсяПриДругомТексте()
    {
        const long newLoadNumber = 44600008;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613", "2421-051 Шапки_отгрузка Номер загрузки " + newLoadNumber, 2000),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

        Assert.Empty(update.PlannedMatches);
        Assert.Single(update.NewRows);
    }

    [Fact]
    public void СтрокаЗаказыБудутЗагружены_ОстаётсяЕслиПоставкаНеПришла()
    {
        const long newLoadNumber = 44600006;
        var update = Build(
            new[]
            {
                Order(2, "Казань-М613", "2446-001 ШПП_отгрузка по готовности Номер загрузки " + newLoadNumber, 100),
            },
            new[] { Journal(0, newLoadNumber, "ЗАПУЩЕН") });

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
                Order(5, "Ozon", "Озон Екб МЗ2446-002 (Шапки) РФ", 13191),
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
