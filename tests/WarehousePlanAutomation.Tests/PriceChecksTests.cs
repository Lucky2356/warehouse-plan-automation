using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class PriceChecksTests
{
    private static PriceRowValues Row(
        PriceListRow? reference = null,
        MarkupRule? markup = null,
        bool duplicate = false) =>
        new(0, 0, "1431280202", "431-329", 1010d, 6.55d, 6615.5d, 6.55d,
            reference ?? Reference(), string.Empty, markup ?? Markup(), duplicate);

    private static PriceListRow Reference() => new(
        "1431280202", "155588", "SLEEP MASK", "МЕЛКИЕ АКСЕССУАРЫ", "РАЗНОЕ",
        "МАСКА ДЛЯ СНА", "431-2802", "СЕРЫЙ-ЧЕРНЫЙ", "WINTER 26-27", "TU");

    private static MarkupRule Markup() => new("МЕЛКИЕ АКСЕССУАРЫ", "ПРОЧЕЕ", 5.24d, 4.14d);

    private static PriceRowState State(
        object? storePeriod = null,
        object? wallAnalogue = null,
        object? googleAnalogue = null,
        object? priceCheck = null,
        IReadOnlyList<string>? grades = null) =>
        new(
            0,
            "431-2802СЕРЫЙ-ЧЕРНЫЙTU",
            storePeriod ?? "01/08-31/01",
            0d,
            wallAnalogue ?? 0d,
            googleAnalogue ?? "нет",
            priceCheck ?? true,
            grades ?? PriceSchema.Prices.AllGrades);

    /// <summary>По три магазина на грейд; <paramref name="denied"/> - «грейд:запрещено магазинов».</summary>
    private static LinkEntry Link(
        string grades = "A,B,C1,C2,D,E,F,G,O120,O140",
        DateTime? start = null,
        DateTime? end = null,
        params string[] denied)
    {
        var deniedStores = denied
            .Select(item => item.Split(':'))
            .ToDictionary(parts => parts[0], parts => int.Parse(parts[1]));

        var list = grades
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(grade => new LinkGrade(grade, 3, 3 - deniedStores.GetValueOrDefault(grade)))
            .ToList();

        return new(
            "431-2802серый-черныйtu",
            list,
            (start ?? new DateTime(2026, 8, 1)).ToOADate(),
            (end ?? new DateTime(2027, 1, 31)).ToOADate(),
            list.Sum(grade => grade.Stores),
            list.Sum(grade => grade.AllowedStores));
    }

    [Fact]
    public void ВсёСошлось_ПишетОк()
    {
        // Вкл везде 1, Deny goods везде 0, даты совпали, грейды сходятся.
        var check = PriceChecks.Check(Row(), State(), Link());

        Assert.Equal(PriceChecks.LinkPresent, check.LinkValue);
        Assert.Equal(PriceChecks.BanCheckOk, check.BanCheck);
        Assert.Empty(check.Highlight);
        Assert.Empty(check.Warnings);
    }

    [Fact]
    public void ОстровныеГрейдыОтмеченыИРазрешены_НеРасхождение()
    {
        // Скрины 1-2: для островов в «link» всё включено, в «Цены» O120 и O140 стоят 1.
        // Раньше программа этих колонок не видела и писала «разрешены, но не отмечены».
        var check = PriceChecks.Check(
            Row(),
            State() with { GradeColumns = PriceSchema.Prices.AllGrades },
            Link(grades: "A,B,C1,C2,D,E,F,G,О120,О140"));

        Assert.Equal(PriceChecks.BanCheckOk, check.BanCheck);
    }

    [Fact]
    public void ГрейдЗапрещёнВоВсехМагазинах_ПишетЗапретНаГрейдах()
    {
        // Скрины 3-4: у F и G «Deny goods подразделение» = 1 во всех магазинах.
        var check = PriceChecks.Check(Row(), State(), Link(denied: new[] { "F:3", "G:3" }));

        Assert.Contains("Запрет на грейдах F, G", check.BanCheck);
        Assert.Contains("F", check.Highlight);
        Assert.Contains("G", check.Highlight);
        Assert.Contains(check.Warnings, warning => warning.Message.Contains("Запрет на грейдах F, G"));
        Assert.NotEqual(PriceChecks.BanCheckOk, check.BanCheck);
    }

    [Theory]
    [InlineData("O120", "O140")]
    [InlineData("О120", "О140")]
    public void ОстроваНеОтмеченыИВЛинкеВыключены_Ок(string island120, string island140)
    {
        // Скрин: «Период продаж острова», O120 и O140 = 0, в «link» у островов Вкл = 0 или
        // Deny goods = 1. АЦР на острова не идёт - так и должно быть, это «Ок».
        var shops = new[] { "A", "B", "C2", "C1", "D", "E", "F", "G" };
        var check = PriceChecks.Check(
            Row(),
            State(grades: shops) with { GradeColumns = PriceSchema.Prices.AllGrades },
            Link(grades: "A,B,C1,C2,D,E,F,G," + island120 + "," + island140,
                denied: new[] { island120 + ":3", island140 + ":3" }));

        Assert.Equal(PriceChecks.BanCheckOk, check.BanCheck);
        Assert.Empty(check.Highlight);
        Assert.Empty(check.Warnings);
    }

    [Fact]
    public void ОстроваОтмеченыНоВЛинкеВыключены_ЗапретНаГрейдах()
    {
        // Обратный случай: O120 и O140 в «Цены» стоят 1 - запрет в «link» противоречит.
        var check = PriceChecks.Check(
            Row(),
            State() with { GradeColumns = PriceSchema.Prices.AllGrades },
            Link(denied: new[] { "O120:3", "O140:3" }));

        Assert.Contains("Запрет на грейдах O120, O140", check.BanCheck);
        Assert.Contains("O120", check.Highlight);
        Assert.NotEqual(PriceChecks.BanCheckOk, check.BanCheck);
    }

    [Fact]
    public void ЗапретВЧастиМагазинов_ВКолонкеНоНеВЗамечаниях()
    {
        var check = PriceChecks.Check(Row(), State(), Link(denied: new[] { "E:2" }));

        Assert.Contains("Запрет в части магазинов", check.BanCheck);
        Assert.Contains("E - 2 из 3", check.BanCheck);
        Assert.Empty(check.Warnings);
        Assert.NotEqual(PriceChecks.BanCheckOk, check.BanCheck);
    }

    [Fact]
    public void ДатыНеСовпали_НеОк()
    {
        var check = PriceChecks.Check(Row(), State(storePeriod: "01/10-31/01"), Link());

        Assert.NotEqual(PriceChecks.BanCheckOk, check.BanCheck);
    }

    [Fact]
    public void КолонкиГрейдаНаЛистеНет_ГрейдНеСверяется()
    {
        var columns = new[] { "A", "B", "C1", "C2", "D", "E", "F", "G" };
        var check = PriceChecks.Check(
            Row(),
            State(grades: columns) with { GradeColumns = columns },
            Link());

        Assert.Equal(PriceChecks.BanCheckOk, check.BanCheck);
    }

    [Fact]
    public void АЦРНетВЛинке_ПишетЧтоНужноДобавить()
    {
        var check = PriceChecks.Check(Row(), State(), link: null);

        Assert.Equal(PriceChecks.LinkMissing, check.LinkValue);
        Assert.Contains("добавить в Линк", check.BanCheck);
        Assert.Contains(PriceSchema.Prices.Link, check.Highlight);
    }

    [Fact]
    public void ДатыНеСовпали_ПишетОбеПары()
    {
        var check = PriceChecks.Check(
            Row(),
            State(storePeriod: "01/10-31/01"),
            Link(start: new DateTime(2026, 8, 1), end: new DateTime(2027, 1, 31)));

        Assert.Contains("01/08-31/01", check.BanCheck);
        Assert.Contains("01/10-31/01", check.BanCheck);
        Assert.Contains(PriceSchema.Prices.StoreSalesPeriod, check.Highlight);
    }

    [Fact]
    public void ГрейдыРасходятся_ПишетЧтоЛишнееИЧегоНеХватает()
    {
        var check = PriceChecks.Check(
            Row(),
            State(grades: new[] { "A", "B", "D" }),
            Link(grades: "A,B,C1"));

        Assert.Contains("магазинов этих грейдов нет: D", check.BanCheck);
        Assert.Contains("не отмечены: C1", check.BanCheck);
        Assert.Contains("D", check.Highlight);
    }

    [Fact]
    public void ПериодПродажНеПодтянулся_ПодсвечиваетИПишетПроСтенки()
    {
        var check = PriceChecks.Check(
            Row(),
            State(storePeriod: CellError.NotAvailable),
            Link());

        Assert.Contains(PriceSchema.Prices.StoreSalesPeriod, check.Highlight);
        Assert.Contains(check.Warnings, w => w.Message.Contains("стенк"));
    }

    [Fact]
    public void РазорваннаяСсылкаНаСтенки_ОтличаетсяОтОтсутствияВСтенках()
    {
        // «#ССЫЛКА!» значит, что сетевая папка со стенками была недоступна, а не что
        // позиции там нет. Причина другая - и подсказка другая.
        var check = PriceChecks.Check(
            Row(),
            State(storePeriod: CellError.Reference),
            Link());

        Assert.Contains(check.Warnings, w => w.Message.Contains("сетевая папка"));
        Assert.DoesNotContain(check.Warnings, w => w.Message.Contains("в стенках нет"));
    }

    [Fact]
    public void РазорваннаяСсылкаВКолонкеОстровов_ТожеЗамечается()
    {
        var state = State() with { IslandPeriod = CellError.Reference };
        var check = PriceChecks.Check(Row(), state, Link());

        Assert.Contains(check.Warnings, w => w.Message.Contains("сетевая папка"));
    }

    [Fact]
    public void ПериодПродажРавенНулю_ПодсвечиваетИПишет()
    {
        var check = PriceChecks.Check(Row(), State(storePeriod: 0d), Link());

        Assert.Contains(PriceSchema.Prices.StoreSalesPeriod, check.Highlight);
        Assert.Contains(check.Warnings, w => w.Message.Contains("нулю"));
    }

    [Theory]
    [InlineData(0d, "2445-136СЕРЫЙ")]
    [InlineData("2445-136СЕРЫЙ", "нет")]
    public void АналогиНеСходятся_Подсвечивает(object wall, object google)
    {
        var check = PriceChecks.Check(
            Row(),
            State(wallAnalogue: wall, googleAnalogue: google),
            Link());

        Assert.Contains(PriceSchema.Prices.GoogleAnalogue, check.Highlight);
        Assert.Contains(check.Warnings, w => w.Message.Contains("Аналоги не сходятся"));
    }

    [Theory]
    [InlineData(0d, "нет")]
    [InlineData("2445-136СЕРЫЙ", "2445-136СЕРЫЙ")]
    public void АналогиСходятся_НетПометки(object wall, object google)
    {
        var check = PriceChecks.Check(
            Row(),
            State(wallAnalogue: wall, googleAnalogue: google),
            Link());

        Assert.DoesNotContain(PriceSchema.Prices.GoogleAnalogue, check.Highlight);
    }

    [Fact]
    public void ПроверкаЦеныЛожь_Подсвечивает()
    {
        var check = PriceChecks.Check(Row(), State(priceCheck: false), Link());

        Assert.Contains(PriceSchema.Prices.PriceCheck, check.Highlight);
        Assert.Contains(check.Warnings, w => w.Message.Contains("Согласованная цена"));
    }

    [Fact]
    public void ПроверкаЦеныНеПосчиталась_Подсвечивает()
    {
        var check = PriceChecks.Check(Row(), State(priceCheck: CellError.NotAvailable), Link());

        Assert.Contains(PriceSchema.Prices.PriceCheck, check.Highlight);
    }

    [Fact]
    public void ШтрихкодНеНайденВПрайсе_ПодсвечиваетКопиюШК()
    {
        var check = PriceChecks.Check(
            Row() with { Reference = null },
            State(),
            Link());

        Assert.Contains(PriceSchema.Prices.BarcodeCopy, check.Highlight);
    }

    [Fact]
    public void ПовторШтрихкода_ПодсвечиваетЦенуЗакупочную()
    {
        var check = PriceChecks.Check(Row(duplicate: true), State(), Link());

        Assert.Contains(PriceSchema.Prices.PurchasePrice, check.Highlight);
    }

    [Fact]
    public void НаценкаНеВыбрана_ПодсвечиваетКолонкуНаценки()
    {
        var check = PriceChecks.Check(Row() with { Markup = null }, State(), Link());

        Assert.Contains(PriceSchema.Prices.PlannedMarkup, check.Highlight);
    }

    [Fact]
    public void ВЛинкеВсеСтрокиЗапрещены_ПишетОбЭтом()
    {
        var check = PriceChecks.Check(Row(), State(), Link(grades: "A,B", denied: new[] { "A:3", "B:3" }));

        Assert.Contains("Во всех магазинах «link» этот АЦР запрещён", check.BanCheck);
    }

    [Fact]
    public void СводкаLink_СчитаетМагазиныПоГрейдуБезРазницыЛатиницыИКириллицы()
    {
        var rows = new[]
        {
            new LinkRow("acr", 0d, 0d, 1d, "О120", 46296d, 46418d),
            new LinkRow("acr", 1d, 0d, 1d, "O120", 46296d, 46418d),
            new LinkRow("acr", 1d, 0d, 1d, "F", 46296d, 46418d),
            new LinkRow("acr", 0d, 0d, null, "G", null, null),
            new LinkRow("acr", 0d, 0d, 0d, "G", null, null),
        };

        var entry = LinkSheetReader.Summarize("acr", rows);

        var island = Assert.Single(entry.Grades, grade => LinkSheetReader.GradeKey(grade.Grade) == "O120");
        Assert.Equal(2, island.Stores);
        Assert.Equal(1, island.AllowedStores);
        Assert.True(Assert.Single(entry.Grades, grade => grade.Grade == "F").IsFullyDenied);
        Assert.True(Assert.Single(entry.Grades, grade => grade.Grade == "G").IsPartlyDenied);
        Assert.Equal(5, entry.Rows);
        Assert.Equal(2, entry.AllowedRows);
        Assert.Equal(46296d, entry.DateStart);
    }

    /// <summary>
    /// «link» как в примере распреда: у островов «Группа_ам» пустая, «Концепт» - ОСТРОВ.
    /// Магазины A..G разрешены, острова - все с «Deny goods» = 1, если <paramref name="islandsDenied"/>.
    /// </summary>
    private static LinkEntry LinkWithConceptIslands(bool islandsDenied)
    {
        var rows = new List<LinkRow>();
        foreach (var grade in new[] { "A", "B", "C1", "C2", "D", "E", "F", "G" })
        {
            rows.Add(new LinkRow("acr", 0d, 0d, 1d, grade, 46235d, 46418d, "МАГАЗИН"));
        }

        for (var i = 0; i < 36; i++)
        {
            rows.Add(new LinkRow("acr", islandsDenied ? 1d : 0d, 0d, 1d, "", 46235d, 46418d, "ОСТРОВ"));
        }

        return LinkSheetReader.Summarize("acr", rows);
    }

    [Fact]
    public void ОстроваБезГрейдаПоКонцепту_НеОтмеченыИЗапрещены_Ок()
    {
        // Пример распреда, 431-2802: O120 и O140 = 0, в «link» все 36 островов запрещены.
        // Раньше: «Запрет в магазинах без «Группа_ам»: 36 из 36» и не «Ок».
        var shops = new[] { "A", "B", "C2", "C1", "D", "E", "F", "G" };
        var check = PriceChecks.Check(
            Row(),
            State(grades: shops) with { GradeColumns = PriceSchema.Prices.AllGrades },
            LinkWithConceptIslands(islandsDenied: true));

        Assert.Equal(PriceChecks.BanCheckOk, check.BanCheck);
        Assert.Empty(check.Warnings);
    }

    [Fact]
    public void ОстроваБезГрейдаПоКонцепту_ОтмеченыНоЗапрещены_Замечание()
    {
        var check = PriceChecks.Check(
            Row(),
            State() with { GradeColumns = PriceSchema.Prices.AllGrades },
            LinkWithConceptIslands(islandsDenied: true));

        Assert.Contains(PriceSchema.Link.IslandsWithoutGrade, check.BanCheck);
        Assert.NotEmpty(check.Warnings);
    }

    [Fact]
    public void ОстроваБезГрейдаПоКонцепту_НеОтмеченыИРазрешены_Ок()
    {
        // Примеры 431-3339 и 431-3340: острова включены, O120 и O140 = 0 - как и раньше, без замечаний.
        var shops = new[] { "A", "B", "C2", "C1", "D", "E", "F", "G" };
        var check = PriceChecks.Check(
            Row(),
            State(grades: shops) with { GradeColumns = PriceSchema.Prices.AllGrades },
            LinkWithConceptIslands(islandsDenied: false));

        Assert.Equal(PriceChecks.BanCheckOk, check.BanCheck);
    }

    [Fact]
    public void МПРавноКоличествуПоставки_ПодсвечиваетИПросятПроверитьСлежение()
    {
        var check = PriceChecks.Check(Row(), State() with { Marketplace = 1010d }, Link());

        Assert.Contains(PriceSchema.Prices.Marketplace, check.Highlight);
        Assert.Contains(check.Warnings, warning => warning.Message.Contains("поставка МП"));
    }

    [Theory]
    [InlineData(500d)]
    [InlineData(0d)]
    [InlineData(null)]
    public void МПНеРавноКоличеству_НичегоНеПишет(double? marketplace)
    {
        var check = PriceChecks.Check(Row(), State() with { Marketplace = marketplace }, Link());

        Assert.DoesNotContain(PriceSchema.Prices.Marketplace, check.Highlight);
        Assert.Empty(check.Warnings);
    }

    [Fact]
    public void МПСОшибкойСтенок_НеСчитаетсяСовпадением()
    {
        var check = PriceChecks.Check(Row(), State() with { Marketplace = CellError.NotAvailable }, Link());

        Assert.DoesNotContain(PriceSchema.Prices.Marketplace, check.Highlight);
    }
}
