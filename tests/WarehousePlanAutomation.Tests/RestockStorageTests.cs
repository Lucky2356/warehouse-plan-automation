using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class StoragePlacesTests
{
    [Theory]
    [InlineData("Маркетплейс", StoragePlace.Marketplace)]
    [InlineData("Хранение", StoragePlace.Storage)]
    [InlineData("Хранилище", StoragePlace.Storage)]
    [InlineData("Возвраты", StoragePlace.Returns)]
    [InlineData("Времянка", StoragePlace.Returns)]
    [InlineData("ОЛД", StoragePlace.Returns)]
    [InlineData("Олды", StoragePlace.Returns)]
    public void ТипХранения(string type, StoragePlace expected) =>
        Assert.Equal(expected, StoragePlaces.FromStorageType(type));

    [Theory]
    [InlineData("Образцы")]
    [InlineData("Брак уценка")]
    [InlineData("Нет Маркировки")]
    [InlineData("")]
    public void ТипыНеДляПодтоварки(string type) => Assert.Null(StoragePlaces.FromStorageType(type));

    [Theory]
    [InlineData("М318-154", StoragePlace.Supplies)]
    [InlineData("Л369-063", StoragePlace.Supplies)]
    [InlineData("С357-039", StoragePlace.NetworkSupplies)]
    [InlineData("C357-039", StoragePlace.NetworkSupplies)]
    public void НомерПоставки(string number, StoragePlace expected) =>
        Assert.Equal(expected, StoragePlaces.FromSupplyNumber(number));

    [Fact]
    public void КолонкиИЛисты()
    {
        Assert.Equal("1МП", StoragePlaces.LoadColumn(StoragePlace.Marketplace));
        Assert.Equal("2МПП", StoragePlaces.LoadColumn(StoragePlace.Supplies));
        Assert.Equal("3А", StoragePlaces.LoadColumn(StoragePlace.Storage));
        Assert.Equal("4СЗП", StoragePlaces.LoadColumn(StoragePlace.NetworkSupplies));
        Assert.Equal("5В", StoragePlaces.LoadColumn(StoragePlace.Returns));
        Assert.Equal("измпп", StoragePlaces.PickSheet(StoragePlace.Supplies));
        Assert.Equal("иза", StoragePlaces.PickSheet(StoragePlace.Storage));
    }
}

public class ReserveTargetsTests
{
    [Theory]
    [InlineData("Золотое Яблоко Сумки из адресов МП, А1,А2,А3, Возвратов приоритет к 04.09", StoragePlace.Marketplace)]
    [InlineData("Ozon МСК Микс из А1,А2,А3 приоритет к 31.08", StoragePlace.Storage)]
    [InlineData("Срочная подтоварка 09.09_Хранение, хранилище Номер загрузки 44814052", StoragePlace.Storage)]
    [InlineData("Заказ интерент магазина № 45297763-0493-1", StoragePlace.Storage)]
    [InlineData("Ozon Мск Микс из Возвратов приоритет к 17.08", StoragePlace.Returns)]
    [InlineData("Пуховики ликвиды FW26-27 осн мерч_из возвратов, времянки_получение в рознице 20.09", StoragePlace.Returns)]
    [InlineData("Lamoda Тапочки из М369-069, Л369-063 приоритет к 07.09.2026", StoragePlace.Supplies)]
    [InlineData("Ламода МСК МЗ352-133, МЗ326-061 (Шапки) FTL 72", StoragePlace.Supplies)]
    [InlineData("Мск Озон СЗ327-046 (Носки) LTL 198 Приоритет К 13.09 микс", StoragePlace.NetworkSupplies)]
    [InlineData("WB Новосибирск СЦ Подтоварка из С326-058, С352-128 приоритет к 31.08", StoragePlace.NetworkSupplies)]
    [InlineData("Зимняя обувь_ликвиды FW26-27_с хранилища, возвратов, времянки_в рознице с 01.10", StoragePlace.Storage)]
    public void МестоРезерва_ПервоеНазванное(string comment, StoragePlace expected)
    {
        var target = ReserveTargets.Parse(comment);

        Assert.Equal(ReserveKind.Place, target.Kind);
        Assert.Equal(expected, target.Place);
    }

    [Theory]
    [InlineData("Опт, обувь, деми 2я часть (от М-Логистик на ООО МаксимаГрупп).", ReserveKind.Anywhere)]
    [InlineData("на образцы, хранение на складе СЗ2437-029 (кеды)", ReserveKind.Samples)]
    [InlineData("Пакеты сентябрь от 07.09", ReserveKind.Unknown)]
    [InlineData("2437-032 Обувь МОНО_в рознице с 01.10", ReserveKind.Unknown)]
    public void ОсобыеРезервы(string comment, ReserveKind expected) =>
        Assert.Equal(expected, ReserveTargets.Parse(comment).Kind);
}

public class StorageAllocatorTests
{
    private static Dictionary<StoragePlace, double> Stock(
        double mp = 0, double mpp = 0, double a = 0, double szp = 0, double v = 0) => new()
    {
        [StoragePlace.Marketplace] = mp,
        [StoragePlace.Supplies] = mpp,
        [StoragePlace.Storage] = a,
        [StoragePlace.NetworkSupplies] = szp,
        [StoragePlace.Returns] = v,
    };

    private static ReserveLine[] Reserve(double quantity, string comment) => new[] { new ReserveLine(quantity, comment) };

    [Fact]
    public void ОдноМесто_ПоОчередиМпМппАСзпВ()
    {
        // МП не хватает, МПП хватает с запасом - берём с МПП, хотя на «А» товара больше.
        var allocation = StorageAllocator.Allocate(50, Stock(mp: 45, mpp: 1380, a: 2584), Array.Empty<ReserveLine>());

        Assert.Equal("МПП", allocation.Text);
        Assert.True(allocation.SinglePlace);
    }

    [Fact]
    public void ОдноМесто_ВсеРезервыАцрВычитаются()
    {
        // 5 на МП минус 10 резерва не хватит, а на «А» 72 - 10 - 5 = 57.
        var allocation = StorageAllocator.Allocate(5, Stock(mp: 5, a: 72, v: 7), Reserve(10, "Пуховики_из возвратов, времянки"));

        Assert.Equal("А", allocation.Text);
    }

    [Fact]
    public void ПримерИнструкции_НаОбразцы_РезервНеВычитается()
    {
        var allocation = StorageAllocator.Allocate(
            15, Stock(mp: 10, a: 5), Reserve(1, "на образцы, хранение на складе СЗ2437-029 (кеды)"));

        Assert.Equal("МП10, А5", allocation.Text);
        Assert.True(allocation.Complete);
        Assert.False(allocation.SinglePlace);
    }

    [Fact]
    public void ПримерИнструкции_РезервМп_ВычитаетсяИзМп()
    {
        var allocation = StorageAllocator.Allocate(
            15, Stock(mp: 10, a: 6),
            Reserve(1, "Золотое Яблоко Кеды из адресов МП, А1,А2,А3, Возвратов приоритет к 04.09"));

        Assert.Equal("МП9, А6", allocation.Text);
        Assert.True(allocation.Complete);
    }

    [Fact]
    public void НеХватило_ЦифраСколькоЕсть()
    {
        var allocation = StorageAllocator.Allocate(7, Stock(a: 7), Reserve(1, "Заказ интерент магазина № Т143437"));

        Assert.Equal("А6", allocation.Text);
        Assert.False(allocation.Complete);
        Assert.Equal(1, allocation.Missing);
        Assert.Equal(6, allocation.Collected);
    }

    [Fact]
    public void НетНигде_Ноль()
    {
        var allocation = StorageAllocator.Allocate(3, Stock(szp: -1), Array.Empty<ReserveLine>());

        Assert.Equal("0", allocation.Text);
        Assert.Equal(3, allocation.Missing);
    }

    [Fact]
    public void Опт_МожетВзятьОткудаУгодно()
    {
        var allocation = StorageAllocator.Allocate(2, Stock(v: 2), Reserve(26, "Опт, головные уборы, палантины"));

        Assert.Equal("В", allocation.Text);
        Assert.True(allocation.SinglePlace);
    }

    [Fact]
    public void НепонятныйРезерв_НеВычитаетсяИЗапоминается()
    {
        var allocation = StorageAllocator.Allocate(4, Stock(mp: 2, a: 3), Reserve(2, "Пакеты сентябрь от 07.09"));

        Assert.Equal("МП2, А2", allocation.Text);
        Assert.Single(allocation.UnknownReserves);
    }
}

public class PickListBuilderTests
{
    [Fact]
    public void ПервыйКод_Хватает()
    {
        var lines = PickListBuilder.Build(
            0, StoragePlace.Marketplace, 5, new[] { new StockCode("154189", 10, ""), new StockCode("1", 30, "") });

        var line = Assert.Single(lines);
        Assert.Equal("154189", line.Code);
        Assert.Equal(PickMark.None, line.Mark);
    }

    [Fact]
    public void ДругойКод_ГдеХватает_Голубой()
    {
        // «хотели взять 5, на первом коде 1, есть коды с 3, 1 и 30 - берём код, где 30».
        var lines = PickListBuilder.Build(
            0, StoragePlace.Storage, 5,
            new[] { new StockCode("a", 1, ""), new StockCode("b", 3, ""), new StockCode("c", 30, "") });

        var line = Assert.Single(lines);
        Assert.Equal("c", line.Code);
        Assert.Equal(PickMark.ReplacedCode, line.Mark);
    }

    [Fact]
    public void НесколькоКодов_Зелёные()
    {
        // «на одном коде 4 шт, на другом 1 шт» - строка размножается.
        var lines = PickListBuilder.Build(
            0, StoragePlace.Returns, 5, new[] { new StockCode("a", 1, ""), new StockCode("b", 4, "") });

        Assert.Equal(2, lines.Count);
        Assert.All(lines, line => Assert.Equal(PickMark.SplitCode, line.Mark));
        Assert.Equal(("b", 4d), (lines[0].Code, lines[0].Quantity));
        Assert.Equal(("a", 1d), (lines[1].Code, lines[1].Quantity));
        Assert.All(lines, line => Assert.Equal(0d, line.Shortage));
    }

    [Fact]
    public void НаКодахНеХватает_Нехватка()
    {
        var lines = PickListBuilder.Build(
            0, StoragePlace.Supplies, 6, new[] { new StockCode("a", 2, "М318-154"), new StockCode("b", 3, "М318-157") });

        Assert.Equal(1d, lines.Sum(line => line.Shortage));
        Assert.Equal(6d, lines.Sum(line => line.Quantity));
    }
}

public class RestockLoaderBuilderTests
{
    [Theory]
    [InlineData(StoragePlace.Marketplace, "любой", "ЗМП-любой")]
    [InlineData(StoragePlace.Storage, "мелкий товар", "ЗА-мелкий")]
    [InlineData(StoragePlace.Marketplace, "ОЧКИ - отдельный заказ", "ЗМП-очки")]
    [InlineData(StoragePlace.Supplies, "352-148ШОКОЛАДНЫЙ - отдельный МОНО заказ", "ЗМПП-352")]
    [InlineData(StoragePlace.Returns, "", "ЗВ")]
    public void НазваниеЛиста(StoragePlace place, string comment, string expected) =>
        Assert.Equal(expected, RestockLoaderBuilder.SheetName(place, comment, new HashSet<string>()));

    [Fact]
    public void НазваниеЛиста_ПовторПолучаетНомер()
    {
        var taken = new HashSet<string>();

        Assert.Equal("ЗМП-352", RestockLoaderBuilder.SheetName(StoragePlace.Marketplace, "352-148ЧЕРНЫЙ - отдельно", taken));
        Assert.Equal("ЗМП-352 2", RestockLoaderBuilder.SheetName(StoragePlace.Marketplace, "352-141 отдельно", taken));
    }

    [Fact]
    public void КаждыйКоментСвойЗагрузочник()
    {
        var comments = new[] { "любой", "Очки отдельно", "любой" };
        var lines = new[]
        {
            new PickLine(0, StoragePlace.Storage, "1", 1, 1, "", PickMark.None),
            new PickLine(1, StoragePlace.Storage, "2", 1, 1, "", PickMark.None),
            new PickLine(1, StoragePlace.Marketplace, "3", 1, 1, "", PickMark.None),
            new PickLine(2, StoragePlace.Storage, "4", 1, 1, "", PickMark.None),
        };

        var loaders = RestockLoaderBuilder.Build(lines, index => comments[index]);

        Assert.Equal(new[] { "ЗМП-очки", "ЗА-любой", "ЗА-очки" }, loaders.Select(l => l.SheetName));
        Assert.Equal(2, loaders[1].Lines.Count);
    }

    [Fact]
    public void КомментарийЗагрузочника()
    {
        var text = RestockLoaderBuilder.CommentText(
            "Lamoda",
            new[] { "СУМКИ", "ОБУВЬ", "СУМКИ" },
            StoragePlace.NetworkSupplies,
            new[] { "С318-156", "С2437-027", "С318-156" },
            new[] { "к 17.09", "к 17.09" });

        Assert.Equal("Lamoda Подтоварка Сумки, Обувь из С318-156, С2437-027 Приоритет к 17.09", text);
    }

    [Fact]
    public void Приоритет_ТекстИлиДата()
    {
        Assert.Equal("к 17.09", RestockLoaderBuilder.PriorityText("к 17.09"));
        Assert.Equal("к 07.09.2026", RestockLoaderBuilder.PriorityText(new DateTime(2026, 9, 7).ToOADate()));
        Assert.Equal(string.Empty, RestockLoaderBuilder.PriorityText(null));
    }

    [Theory]
    [InlineData(StoragePlace.Marketplace, "адресов МП")]
    [InlineData(StoragePlace.Storage, "А1,А2,А3")]
    [InlineData(StoragePlace.Returns, "Возвратов")]
    public void МестоВКомментарии(StoragePlace place, string expected) =>
        Assert.Equal(expected, RestockLoaderBuilder.PlaceText(place, Array.Empty<string>()));

    [Theory]
    [InlineData("208", null, "Lamoda")]
    [InlineData("206", null, "Ozon")]
    [InlineData("211", null, "Золотое Яблоко")]
    [InlineData("300", "МеркетПлейсы / Сбермегамаркет", "Сбермегамаркет")]
    public void Маркетплейс(string code, string? group, string expected) =>
        Assert.Equal(expected, RestockLoaderBuilder.MarketplaceName(code, group));
}

public class ReservePeriodTests
{
    private static readonly DateTime September = new(2026, 9, 14);

    [Theory]
    [InlineData(2026, 7, 1, true)]
    [InlineData(2026, 9, 30, true)]
    [InlineData(2026, 6, 30, false)]
    [InlineData(2025, 9, 14, false)]
    public void СентябрьОставляетИюльАвгустСентябрь(int year, int month, int day, bool keep) =>
        Assert.Equal(keep, ReservePeriod.Keep(new DateTime(year, month, day), September));

    [Fact]
    public void ВЯнвареПрошлыйГодУдаляется() =>
        Assert.False(ReservePeriod.Keep(new DateTime(2026, 12, 20), new DateTime(2027, 1, 10)));
}
