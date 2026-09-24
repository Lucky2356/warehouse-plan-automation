using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
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
    public void ТипХранения(string type, StoragePlace expected) =>
        Assert.Equal(expected, StoragePlaces.FromStorageType(type));

    [Theory]
    [InlineData("Образцы")]
    [InlineData("Брак уценка")]
    [InlineData("Нет Маркировки")]
    [InlineData("ОЛД")]
    [InlineData("Олды")]
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
        // Цифра - очередь, в которой место берётся: МП, МПП, В, А, СЗП.
        Assert.Equal("1МП", StoragePlaces.LoadColumn(StoragePlace.Marketplace));
        Assert.Equal("2МПП", StoragePlaces.LoadColumn(StoragePlace.Supplies));
        Assert.Equal("3В", StoragePlaces.LoadColumn(StoragePlace.Returns));
        Assert.Equal("4А", StoragePlaces.LoadColumn(StoragePlace.Storage));
        Assert.Equal("5СЗП", StoragePlaces.LoadColumn(StoragePlace.NetworkSupplies));
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
    public void ОдноМесто_ПоОчередиМпМппВАСзп()
    {
        // МП не хватает, МПП хватает с запасом - берём с МПП, хотя на «А» товара больше.
        var allocation = StorageAllocator.Allocate(50, Stock(mp: 45, mpp: 1380, a: 2584), Array.Empty<ReserveLine>());

        Assert.Equal("МПП", allocation.Text);
        Assert.True(allocation.SinglePlace);
    }

    [Fact]
    public void ОдноМесто_ВозвратыРаньшеХранения()
    {
        var allocation = StorageAllocator.Allocate(5, Stock(a: 300, v: 20), Array.Empty<ReserveLine>());

        Assert.Equal("В", allocation.Text);
    }

    [Fact]
    public void ОдноМесто_ВсеРезервыАцрВычитаются()
    {
        // 5 на МП минус 10 резерва не хватит, а на «А» 72 - 10 - 5 = 57.
        var allocation = StorageAllocator.Allocate(5, Stock(mp: 5, a: 72), Reserve(10, "Заказ интерент магазина № Т143437"));

        Assert.Equal("А", allocation.Text);
    }

    [Fact]
    public void НаОбразцы_ВычитаетсяИзПоставок()
    {
        // На образцы и на фото берут с поставок: резерв съедает МПП, и товар идёт с «В».
        var allocation = StorageAllocator.Allocate(
            5, Stock(mpp: 5, v: 5), Reserve(5, "на образцы, хранение на складе СЗ2437-029 (кеды)"));

        Assert.Equal("В", allocation.Text);
    }

    [Fact]
    public void НаОбразцы_БезПоставокНиЧегоНеСъедает()
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
    public void Опт_ВычитаетсяСЛюбогоМеста()
    {
        // Резерв опта в 3 шт. снимается с МП, и до нужных 5 добирают с возвратов.
        var allocation = StorageAllocator.Allocate(5, Stock(mp: 3, v: 4), Reserve(3, "Опт, головные уборы, палантины"));

        Assert.Equal("В4", allocation.Text);
        Assert.Equal(1d, allocation.Missing);
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
    public void НаименьшийКод_Хватает()
    {
        // Код с меньшим номером идёт первым, даже если на листе он стоит ниже.
        var lines = PickListBuilder.Build(
            0, StoragePlace.Marketplace, 5, new[] { new StockCode("154189", 10, ""), new StockCode("151001", 30, "") });

        var line = Assert.Single(lines);
        Assert.Equal("151001", line.Code);
        Assert.Equal(PickMark.None, line.Mark);
    }

    [Fact]
    public void СледующийКод_ГдеХватает_Голубой()
    {
        // На наименьшем коде 1 шт., дальше по возрастанию - 30 и 3; берём ближайший,
        // на котором хватает пяти, а не самый большой.
        var lines = PickListBuilder.Build(
            0, StoragePlace.Storage, 5,
            new[] { new StockCode("154003", 3, ""), new StockCode("154001", 1, ""), new StockCode("154002", 30, "") });

        var line = Assert.Single(lines);
        Assert.Equal("154002", line.Code);
        Assert.Equal(PickMark.ReplacedCode, line.Mark);
    }

    [Fact]
    public void НесколькоКодов_Зелёные()
    {
        // «на одном коде 4 шт, на другом 1 шт» - строка размножается, коды по возрастанию.
        var lines = PickListBuilder.Build(
            0, StoragePlace.Returns, 5, new[] { new StockCode("154002", 4, ""), new StockCode("154001", 1, "") });

        Assert.Equal(2, lines.Count);
        Assert.All(lines, line => Assert.Equal(PickMark.SplitCode, line.Mark));
        Assert.Equal(("154001", 1d), (lines[0].Code, lines[0].Quantity));
        Assert.Equal(("154002", 4d), (lines[1].Code, lines[1].Quantity));
        Assert.All(lines, line => Assert.Equal(0d, line.Shortage));
    }

    [Fact]
    public void НечисловойКод_УходитВКонец()
    {
        var lines = PickListBuilder.Build(
            0, StoragePlace.Storage, 5, new[] { new StockCode("б/н", 30, ""), new StockCode("154001", 10, "") });

        Assert.Equal("154001", Assert.Single(lines).Code);
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
    public void КомментарийЗагрузочника_ПоставкиНазываютсяНомерами()
    {
        var text = RestockLoaderBuilder.CommentText(
            "Lamoda",
            string.Empty,
            "любой",
            new[] { "С318-156", "С2437-027", "С318-156" },
            new[] { "к 17.09", "к 17.09" });

        Assert.Equal("Lamoda Подтоварка Любой из С318-156, С2437-027 Приоритет к 17.09", text);
    }

    [Fact]
    public void КомментарийЗагрузочника_СГородом()
    {
        var text = RestockLoaderBuilder.CommentText(
            "Ozon", "Мск", "микс", Array.Empty<string>(), new[] { "к 28.09" });

        Assert.Equal("Ozon Мск Подтоварка Микс Приоритет к 28.09", text);
    }

    [Fact]
    public void Подразделение206_ЗагрузочникиТолькоПоКоментам()
    {
        var comments = new[] { "микс", "крупное", "микс" };
        var lines = new[]
        {
            new PickLine(0, StoragePlace.Storage, "1", 1, 1, "", PickMark.None),
            new PickLine(1, StoragePlace.Marketplace, "2", 1, 1, "", PickMark.None),
            new PickLine(2, StoragePlace.Returns, "3", 1, 1, "", PickMark.None),
        };

        var loaders = RestockLoaderBuilder.Build(lines, index => comments[index], mergePlaces: true);

        Assert.Equal(new[] { "З-микс", "З-крупное" }, loaders.Select(l => l.SheetName));
        Assert.Equal(2, loaders[0].Lines.Count);
        Assert.Null(loaders[0].Place);
    }

    [Fact]
    public void НомерЗаказа_ПоГородам()
    {
        var loader = new RestockLoader("З-микс", null, "микс", new[]
        {
            new PickLine(0, StoragePlace.Marketplace, "1", 5, 5, "", PickMark.None, 0d, "Мск"),
            new PickLine(0, StoragePlace.Marketplace, "1", 2, 5, "", PickMark.None, 0d, "Спб"),
            new PickLine(0, StoragePlace.Marketplace, "1", 1, 5, "", PickMark.None, 0d, "Екб"),
        });

        Assert.Equal(new[] { 1d, 2d, 3d }, RestockLoaderBuilder.OrderNumbers(loader, new[] { "Мск", "Спб", "Екб" }));
    }

    [Fact]
    public void НомерЗаказа_МоноЗаказКаждыйКодОтдельно()
    {
        var loader = new RestockLoader("З-моно", null, "МОНО заказ", new[]
        {
            new PickLine(0, StoragePlace.Marketplace, "154267", 5, 5, "", PickMark.None),
            new PickLine(1, StoragePlace.Marketplace, "154268", 3, 3, "", PickMark.None),
            new PickLine(2, StoragePlace.Marketplace, "154267", 2, 5, "", PickMark.None),
        });

        Assert.Equal(new[] { 1d, 2d, 1d }, RestockLoaderBuilder.OrderNumbers(loader, Array.Empty<string>()));
    }

    [Fact]
    public void Приоритет_ТекстИлиДата()
    {
        Assert.Equal("к 17.09", RestockLoaderBuilder.PriorityText("к 17.09"));
        Assert.Equal("к 07.09.2026", RestockLoaderBuilder.PriorityText(new DateTime(2026, 9, 7).ToOADate()));
        Assert.Equal(string.Empty, RestockLoaderBuilder.PriorityText(null));
    }

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
    [InlineData(2026, 1, 9, true)]
    [InlineData(2026, 7, 1, true)]
    [InlineData(2026, 9, 30, true)]
    [InlineData(2025, 12, 31, false)]
    [InlineData(2025, 9, 14, false)]
    public void ОстаютсяЗаказыТекущегоГода(int year, int month, int day, bool keep) =>
        Assert.Equal(keep, ReservePeriod.Keep(new DateTime(year, month, day), September));

    [Fact]
    public void ВЯнвареПрошлыйГодУдаляется() =>
        Assert.False(ReservePeriod.Keep(new DateTime(2026, 12, 20), new DateTime(2027, 1, 10)));
}

public class PickListCityTests
{
    private static PickLine Line(string code, double quantity, double shortage = 0d) =>
        new(0, StoragePlace.Marketplace, code, quantity, quantity, string.Empty, PickMark.None, shortage);

    [Fact]
    public void ГородаБерутПоОчереди()
    {
        var lines = PickListBuilder.SplitByCity(
            new[] { Line("a", 6), Line("b", 4) },
            new[] { ("Мск", 5d), ("Спб", 3d), ("Екб", 2d) });

        Assert.Equal(
            new[] { ("a", 5d, "Мск"), ("a", 1d, "Спб"), ("b", 2d, "Спб"), ("b", 2d, "Екб") },
            lines.Select(line => (line.Code, line.Quantity, line.City)));
    }

    [Fact]
    public void ПоследнемуГородуНеХватает()
    {
        // Собрали только 4 из 7: первый город увозит своё, третьему не достаётся ничего.
        var lines = PickListBuilder.SplitByCity(
            new[] { Line("a", 4) }, new[] { ("Мск", 3d), ("Спб", 2d), ("Екб", 2d) });

        Assert.Equal(new[] { ("Мск", 3d), ("Спб", 1d) }, lines.Select(line => (line.City, line.Quantity)));
    }

    [Fact]
    public void ОстатокДостаётсяПоследнемуГороду()
    {
        // Количество по городам меньше собранного - разницу нельзя потерять.
        var lines = PickListBuilder.SplitByCity(new[] { Line("a", 5) }, new[] { ("Мск", 2d) });

        var line = Assert.Single(lines);
        Assert.Equal((5d, "Мск"), (line.Quantity, line.City));
    }

    [Fact]
    public void БезГородов_СтрокиНеМеняются()
    {
        var source = new[] { Line("a", 5) };

        Assert.Same(source, PickListBuilder.SplitByCity(source, Array.Empty<(string, double)>()));
    }
}

public class RestockQuantityTests
{
    private static SheetGrid Headers(params string[] titles) =>
        SheetGrid.FromRows(1, 1, new List<object?[]> { titles.Cast<object?>().ToArray() });

    [Fact]
    public void ОдинГород_ОднаКолонкаВПодтоварку()
    {
        var layout = RestockSchema.Load.ResolveQuantity(Headers("Код", "в подтоварку ", "комент"), 1);

        Assert.Equal(2, layout.Total);
        Assert.Empty(layout.Cities);
    }

    [Fact]
    public void ОдинГород_КолонкаПодписанаСкладом()
    {
        var layout = RestockSchema.Load.ResolveQuantity(Headers("Код", "в подтоварку Мск", "комент"), 1);

        Assert.Equal(2, layout.Total);
        Assert.Empty(layout.Cities);
    }

    [Fact]
    public void НесколькоГородов_СчитаемПоИтогу()
    {
        var layout = RestockSchema.Load.ResolveQuantity(
            Headers("Код", "Склад Мск", "Склад Спб", "Склад Екб", "Итого на МП", "комент"), 1);

        Assert.Equal(5, layout.Total);
        Assert.Equal(new[] { "Мск", "Спб", "Екб" }, layout.Cities.Select(city => city.Name));
        Assert.Equal(new[] { 2, 3, 4 }, layout.Cities.Select(city => city.Column));
    }

    [Fact]
    public void НесколькоГородов_КолонкиВПодтоварку()
    {
        var layout = RestockSchema.Load.ResolveQuantity(
            Headers("Итого в подтоварку", "в подтоварку МСК", "в подтоварку НСК"), 1);

        Assert.Equal(1, layout.Total);
        Assert.Equal(new[] { "МСК", "НСК" }, layout.Cities.Select(city => city.Name));
    }

    [Fact]
    public void НесколькоГородовБезИтога_Останавливаемся()
    {
        var error = Assert.Throws<WorkbookValidationException>(() =>
            RestockSchema.Load.ResolveQuantity(Headers("в подтоварку МСК", "в подтоварку НСК"), 1));

        Assert.Contains("Итого в подтоварку", string.Join(" ", error.Problems));
    }

    [Fact]
    public void КолонкиНет_Останавливаемся() =>
        Assert.Throws<WorkbookValidationException>(() =>
            RestockSchema.Load.ResolveQuantity(Headers("Код", "комент"), 1));
}
