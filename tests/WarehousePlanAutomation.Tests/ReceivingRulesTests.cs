using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

/// <summary>Приемка на хранилище: номера поставок и разбор «Приходов».</summary>
public class SupplyNumbersTests
{
    [Theory]
    [InlineData("С352-136", "352-136", 'с')]
    [InlineData("М318-157", "318-157", 'м')]
    [InlineData("Л318-151", "318-151", 'л')]
    [InlineData("С2412-015", "2412-015", 'с')]
    [InlineData("C352-136", "352-136", 'с')]
    [InlineData("M318-157", "318-157", 'м')]
    public void НомерРазбираетсяНаБуквуИЦифры(string number, string digits, char letter)
    {
        Assert.Equal(digits, SupplyNumbers.Digits(number));
        Assert.Equal(letter, SupplyNumbers.Letter(number));
    }

    [Fact]
    public void СокращённыйПереченьРазворачивается()
    {
        var numbers = SupplyNumbers.Extract(
            "329-081, 431-327,331, 2445-011,014,015,016,017,018 Бижутерия, 331-042 Очки, 322-147 Укр-я для волос");

        Assert.Contains("431-327", numbers);
        Assert.Contains("431-331", numbers);
        Assert.Contains("2445-014", numbers);
        Assert.Contains("2445-018", numbers);
        Assert.Contains("322-147", numbers);
        Assert.DoesNotContain("2445-198", numbers);
        Assert.Equal(11, numbers.Count);
    }

    [Theory]
    [InlineData("2438-003, 327-037,38 Носки_получение в рознице 07.11.2025", "327-038")]
    [InlineData("431-285,86 Брелоки, 431-282, 431-290,Бижутерия_отгрузка по готовности", "431-286")]
    [InlineData("2449-001, 003, 004, 2437-018, 2450-001 Сланцы", "2449-004")]
    [InlineData("2430-026, 028 Обувь СЕТ1_в рознице с 1.10", "2430-028")]
    [InlineData("П338-063, П2453-001 Очки_отгрузка по готовности", "2453-001")]
    public void СокращениеПродолжаетПоследнийПолныйНомер(string text, string expected)
    {
        Assert.Contains(expected, SupplyNumbers.Extract(text));
    }

    [Theory]
    [InlineData("334-030 FTL 38, 314-051 FTL 39, 318-137 FTL 40 Сумки", "334-038")]
    [InlineData("352-136, 2421-049 Шапки, шарфы_в рознице с 07.09", "352-107")]
    [InlineData("Номер загрузки 44231895, 2412-015", "2412-895")]
    public void ЧислаВнеПеречняНеСтановятсяНомерами(string text, string unexpected)
    {
        Assert.DoesNotContain(unexpected, SupplyNumbers.Extract(text));
    }

    [Fact]
    public void НомерНеНаходитсяВнутриДругогоНомера()
    {
        var numbers = SupplyNumbers.Extract("2417-051, 2433-008 Кеды_СЕТ1_получение в рознице 01.08");

        Assert.Contains("2433-008", numbers);
        Assert.DoesNotContain("325-008", numbers);
    }
}

public class IncomingSupplyRulesTests
{
    private static readonly IReadOnlySet<string> Removed = SupplyNumbers.ExtractAll(new[]
    {
        "352-136, 2421-049 Шапки, шарфы_в рознице с 07.09",
        "2445-011,014,015,016,017,018 Бижутерия",
        "2412-015 Пуховики_старый заказ",
    });

    private static readonly IReadOnlySet<string> Planned = SupplyNumbers.ExtractAll(new[]
    {
        "2412-015 Пуховики_получение в рознице 20.09",
    });

    private static IncomingDecision Decide(string number, object? status = null, object? difference = null) =>
        IncomingSupplyRules.Decide(
            new[] { new IncomingRow(0, number, status ?? "ЗАПУЩЕН", difference ?? 520d) },
            Removed,
            Planned).Single();

    [Theory]
    [InlineData("ЗАКРЫТ")]
    [InlineData("Новый")]
    [InlineData("")]
    public void НеЗапущенные_Удаляются(string status)
    {
        Assert.Equal(IncomingKind.Removed, Decide("С352-136", status).Kind);
    }

    [Theory]
    [InlineData(9d, IncomingKind.Removed)]
    [InlineData(0d, IncomingKind.Removed)]
    [InlineData(-4d, IncomingKind.Removed)]
    [InlineData(10d, IncomingKind.Collected)]
    public void РазницаДоДевятиВключительно_Удаляется(double difference, IncomingKind expected)
    {
        Assert.Equal(expected, Decide("С352-136", difference: difference).Kind);
    }

    [Fact]
    public void ЗапущеноСТочкой_ТожеЗапущено()
    {
        Assert.Equal(IncomingKind.Collected, Decide("С352-136", "ЗАПУЩЕНО").Kind);
    }

    [Fact]
    public void ПоставкаНаЛ_Удаляется()
    {
        Assert.Equal(IncomingKind.Removed, Decide("Л318-151").Kind);
    }

    [Fact]
    public void ПоставкаНаМ_Маркетплейс()
    {
        Assert.Equal(IncomingKind.Marketplace, Decide("М318-157").Kind);
    }

    [Fact]
    public void НайденаВУдалённыхИзПлана_Собрана()
    {
        Assert.Equal(IncomingKind.Collected, Decide("С2445-018").Kind);
    }

    [Fact]
    public void НайденаВПланеСклада_НеСобрана_ДажеЕслиЕстьВУдалённых()
    {
        var decision = Decide("С2412-015");

        Assert.Equal(IncomingKind.NotCollected, decision.Kind);
        Assert.Contains("есть и в", decision.Note);
    }

    [Fact]
    public void НигдеНеНайдена_ОстаётсяБезЦвета()
    {
        // Так в шаблоне стоят С2446-004 и С2446-005: аналитик покрасил их сам,
        // а по листам их номера не находятся.
        var decision = Decide("С2446-004");

        Assert.Equal(IncomingKind.Unresolved, decision.Kind);
        Assert.Contains("2446-004", decision.Note);
    }
}

public class ReceivingStockRulesTests
{
    [Fact]
    public void КодТекстомИКодЧисломСовпадают()
    {
        Assert.Equal(ReceivingStockRules.CodeKey(139007d), ReceivingStockRules.CodeKey("139007"));
        Assert.Equal(ReceivingStockRules.CodeKey(31628d), ReceivingStockRules.CodeKey(" 31628"));
    }

    [Fact]
    public void ОстаткиСкладываютсяПоКоду()
    {
        var sums = ReceivingStockRules.SumByCode(new (object?, object?)[]
        {
            ("150242", 29d), ("150242", 1d), ("150243", 23d), ("", 5d), ("150244", "#Н/Д"),
        });

        Assert.Equal(30d, sums["150242"]);
        Assert.Equal(23d, sums["150243"]);
        Assert.Equal(2, sums.Count);
    }

    [Theory]
    [InlineData("13900Б", true)]
    [InlineData("б39007", true)]
    [InlineData("139007", false)]
    public void БракПомеченБуквойБ(string code, bool defect)
    {
        Assert.Equal(defect, ReceivingStockRules.IsDefect(code));
    }

    [Fact]
    public void КодИзАртикула_ПоследниеШестьЗнаков()
    {
        Assert.Equal("139007", ReceivingStockRules.CodeFromArticle("2412-004 139007"));
        Assert.Equal("31628", ReceivingStockRules.CodeFromArticle("000-500 31628"));
    }

    [Fact]
    public void РезервыПрошлыхЛет_Устарели()
    {
        Assert.True(ReceivingStockRules.IsOutdatedReserve(new DateTime(2025, 12, 31).ToOADate(), 2026));
        Assert.False(ReceivingStockRules.IsOutdatedReserve(new DateTime(2026, 2, 11, 11, 2, 0).ToOADate(), 2026));
        Assert.True(ReceivingStockRules.IsOutdatedReserve("11.02.2024 11:02", 2026));
        Assert.False(ReceivingStockRules.IsOutdatedReserve("вчера", 2026));
    }

    [Fact]
    public void ТипРезерва_ТолькоРезервОтгрузка()
    {
        Assert.True(ReceivingStockRules.IsShipmentReserve("РезервОтгрузка"));
        Assert.False(ReceivingStockRules.IsShipmentReserve("РезервПеремещение"));
    }

    [Fact]
    public void ЗапретДляВсех()
    {
        Assert.True(ReceivingStockRules.IsDeniedForAll("все"));
        Assert.True(ReceivingStockRules.IsDeniedForAll(" Все "));
        Assert.False(ReceivingStockRules.IsDeniedForAll("Сибирь"));
    }
}

public class StorageSelectionTests
{
    private static WarehouseRow Row(int index, string type, object? container, object? code, object? quantity) =>
        new(index, type, container, code, quantity, new object?[] { "адрес " + index });

    [Fact]
    public void ОтбираетсяТипХранения_ПовторыТарыУбираются_ПорядокПоКодуИКоличеству()
    {
        var result = StorageSelection.Select(
            new[]
            {
                Row(0, "Хранение", 15161d, 129822d, 97d),
                Row(1, "ОЛД", 15148d, 129822d, 6d),
                Row(2, "Хранение", 15148d, 129822d, 6d),
                Row(3, "Хранение", 15161d, 129822d, 97d),
                Row(4, "Маркетплейс", 999d, 1d, 1d),
                Row(5, "Хранение", 10328d, 31628d, 412d),
                Row(6, "хранение", 15155d, 129822d, 100d),
            },
            ReceivingSchema.Warehouse.StorageTypeStorage);

        Assert.Equal(5, result.Matched);
        Assert.Equal(1, result.DuplicateContainers);
        Assert.Equal(new[] { 5, 2, 0, 6 }, result.Rows.Select(row => row.Index));
    }

    [Fact]
    public void ПустаяТара_ТожеСчитаетсяПовтором()
    {
        var result = StorageSelection.Select(
            new[] { Row(0, "Хранение", null, 1d, 1d), Row(1, "Хранение", "", 2d, 1d) },
            ReceivingSchema.Warehouse.StorageTypeStorage);

        Assert.Single(result.Rows);
        Assert.Equal(1, result.BlankContainers);
    }

    [Fact]
    public void ЧислаРаньшеТекста_ПустыеВКонце()
    {
        var result = StorageSelection.Select(
            new[]
            {
                Row(0, "Хранение", 1d, "АБВ", 1d),
                Row(1, "Хранение", 2d, null, 1d),
                Row(2, "Хранение", 3d, "150649", 1d),
                Row(3, "Хранение", 4d, 31628d, 1d),
            },
            ReceivingSchema.Warehouse.StorageTypeStorage);

        Assert.Equal(new[] { 3, 2, 0, 1 }, result.Rows.Select(row => row.Index));
    }
}

public class ReceivingSummaryTests
{
    private static Dictionary<string, double> Sums(params (string Code, double Value)[] values) =>
        values.ToDictionary(pair => pair.Code, pair => pair.Value, StringComparer.Ordinal);

    [Fact]
    public void ОстаютсяСтрокиСХотяБыОднимИсточником()
    {
        var sources = new SummarySources(
            Sums(("139640", 292d)),
            Sums(("144454", 3d)),
            Sums(("154348", 0d)),
            Sums(("154389", 12d)),
            Sums(("154069", -5d)));

        var kept = ReceivingSummary.SelectRows(
            new object?[] { 139640d, 144454d, 154348d, "154389", 154069d, 999999d, null },
            sources);

        Assert.Equal(new[] { 0, 1, 3 }, kept);
    }

    [Fact]
    public void ОтветМП_ПерваяСтрокаКода_ТолькоБольшеНуля()
    {
        var answers = ReceivingSummary.MarketplaceAnswers(new (object?, object?)[]
        {
            (151170d, 30d), (150590d, 0d), (151170d, 99d), (152602d, "#Н/Д"), (154434d, 100d),
        });

        Assert.Equal(30d, answers["151170"]);
        Assert.Equal(100d, answers["154434"]);
        Assert.Equal(2, answers.Count);
    }

    [Fact]
    public void ВторойЭтап_МестоХраненияИДопоставитьИзМП()
    {
        var fills = ReceivingSummary.PlanFills(new[]
        {
            new SummaryState(0, null, 5d, 0d, 0d),
            new SummaryState(1, 40d, 0d, 0d, 184d),
            new SummaryState(2, 5d, 5d, 10d, 5d),
            new SummaryState(3, null, 0d, 30d, 0d),
            new SummaryState(4, 0d, 0d, 30d, 0d),
            new SummaryState(5, null, 0d, 0d, 12d),
            new SummaryState(6, null, "#Н/Д", -2146826246d, 0d),
        }).ToDictionary(fill => fill.Index);

        Assert.Equal(PlaceSource.StorageAddresses, fills[0].Place);
        Assert.Equal(PlaceSource.CollectedSupply, fills[1].Place);
        Assert.Equal(PlaceSource.CollectedSupply, fills[2].Place);
        Assert.Null(fills[2].RestockFromMarketplace);
        Assert.Equal(PlaceSource.None, fills[3].Place);
        Assert.Equal(30d, fills[3].RestockFromMarketplace);

        // Ноль в «Допоставить» - это решение «не брать», а не пустая ячейка.
        Assert.False(fills.ContainsKey(4));

        // Собранная поставка без решения в «Допоставить» места не даёт.
        Assert.False(fills.ContainsKey(5));
        Assert.False(fills.ContainsKey(6));
    }
}

public class ReceivingFileNameTests
{
    [Fact]
    public void ОтметкиПриемкиСнимаютсяПриСледующемЭтапе()
    {
        var now = new DateTime(2026, 9, 10, 16, 5, 0);
        var prepared = Core.Models.OutputFileName.Build(
            "Приемка на хранилище ШАБЛОН", Core.Models.OutputFileName.ReceivingMark, now);

        Assert.Equal("Приемка на хранилище ШАБЛОН приемка 2026-09-10_1605", prepared);
        Assert.Equal(
            "Приемка на хранилище ШАБЛОН адреса 2026-09-10_1605",
            Core.Models.OutputFileName.Build(prepared, Core.Models.OutputFileName.AddressesMark, now));
    }
}

public class AddressBlocksTests
{
    private static CountedRow Row(string address, double container, double code, double addressNumber, double containerNumber) =>
        new(address, container, code, addressNumber, containerNumber);

    [Fact]
    public void Адреса_ПоКодуОднойСтрокойВПервойСтрокеКода()
    {
        // Пример из инструкции: у 154687 два адреса, у 154680 - один.
        var lines = AddressBlocks.Addresses(new[]
        {
            Row("A3-10.10.02", 18348, 154680, 1, 1),
            Row("A3-09.01.02", 18189, 154687, 1, 1),
            Row("A3-09.02.02", 18166, 154687, 2, 2),
            Row("A3-09.02.02", 18171, 154687, 2, 3),
        });

        Assert.Equal(3, lines.Count);
        Assert.Equal("A3-10.10.02", lines[0].Joined);
        Assert.Equal("A3-09.01.02, A3-09.02.02", lines[1].Joined);
        Assert.Equal(string.Empty, lines[2].Joined);
    }

    [Fact]
    public void Тары_ВсеТарыКодаБезУдаленияПовторовАдреса()
    {
        var lines = AddressBlocks.Containers(new[]
        {
            Row("A3-10.10.02", 18348, 154680, 1, 1),
            Row("A3-09.01.02", 18189, 154687, 1, 1),
            Row("A3-09.02.02", 18166, 154687, 2, 2),
            Row("A3-09.02.02", 18171, 154687, 2, 3),
            Row("A3-09.02.02", 18183, 154687, 2, 4),
            Row("A3-09.02.02", 18184, 154687, 2, 5),
        });

        Assert.Equal(6, lines.Count);
        Assert.Equal("18348", lines[0].Joined);
        Assert.Equal("18189, 18166, 18171, 18183, 18184", lines[1].Joined);
        Assert.All(lines.Skip(2), line => Assert.Equal(string.Empty, line.Joined));
    }
}
