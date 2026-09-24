using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class RestockBanRulesTests
{
    private static RestockRow Row(
        object? ban = null,
        double? quantity = 20d,
        double? available = 100d,
        string sector = "ОБУВЬ",
        string group = "КЕДЫ",
        string article = "369-146") =>
        new(0, sector, group, article, article + "РОЗОВЫЙXS", quantity, ban, available);

    // ===== «Отгрузка в рамках заказа МП, запрет забора из розницы» =====

    private const string OrderOnly = "отгрузка в рамках заказа МП, запрет забора из розницы";

    [Fact]
    public void ЗаказМП_ХватаетНаВБиОзоне_КоличествоНеМеняется()
    {
        var decision = RestockBanRules.Decide(Row(ban: OrderOnly, quantity: 20d, available: 134d));

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.Null(decision.Quantity);
        Assert.False(decision.Highlight);
        Assert.False(decision.NeedsApproval);
    }

    [Fact]
    public void ЗаказМП_ПросятБольшеЧемМожноСобрать_КоличествоУрезается()
    {
        var decision = RestockBanRules.Decide(Row(ban: OrderOnly, quantity: 40d, available: 12d));

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.Equal(12d, decision.Quantity);
        Assert.True(decision.Highlight);
    }

    [Fact]
    public void ЗаказМП_СобратьНечего_ДаётЗамечание()
    {
        var decision = RestockBanRules.Decide(Row(ban: OrderOnly, quantity: 40d, available: 0d));

        Assert.Equal(0d, decision.Quantity);
        Assert.NotNull(decision.Problem);
    }

    // ===== «Запрет забора из розницы» =====

    private const string NoRetail = "Запрет забора из розницы";

    [Fact]
    public void ЗапретРозницы_НаСогласование()
    {
        var decision = RestockBanRules.Decide(Row(ban: NoRetail, sector: "ОБУВЬ"));

        Assert.Equal(RestockSchema.NoteApprove, decision.Note);
        Assert.True(decision.Highlight);
        Assert.True(decision.NeedsApproval);
    }

    [Fact]
    public void ЗапретРозницы_БезЛистаИсключений_СогласовываетсяВсё()
    {
        // Зашитых секторов больше нет: пока лист «Исключения» не заведён, исключений нет.
        var decision = RestockBanRules.Decide(Row(ban: NoRetail, sector: "БИЖУТЕРИЯ"));

        Assert.Equal(RestockSchema.NoteApprove, decision.Note);
        Assert.True(decision.NeedsApproval);
    }

    // ===== Лист «Исключения» =====

    [Theory]
    [InlineData("БИЖУТЕРИЯ")]
    [InlineData("бижутерия")]
    [InlineData("КЕДЫ")]
    [InlineData("369-146")]
    [InlineData("369-146РОЗОВЫЙXS")]
    [InlineData("369-146РОЗОВЫЙ")]
    public void Исключения_ОтдаёмБезСогласования(string value)
    {
        var exceptions = new RestockExceptions(new[] { value });
        var decision = RestockBanRules.Decide(Row(ban: NoRetail, sector: "БИЖУТЕРИЯ"), exceptions);

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.False(decision.NeedsApproval);
    }

    [Fact]
    public void Исключения_ЧужиеЗначенияНеПодходят()
    {
        var exceptions = new RestockExceptions(new[] { "СУМКИ", "800-100" });
        var decision = RestockBanRules.Decide(Row(ban: NoRetail), exceptions);

        Assert.Equal(RestockSchema.NoteApprove, decision.Note);
    }

    [Fact]
    public void Исключения_ДопСоглВсёРавноСогласовывается()
    {
        // «доп согл» проверяется раньше исключений: разговор нужен в любом случае.
        var exceptions = new RestockExceptions(new[] { "ОБУВЬ" });
        var decision = RestockBanRules.Decide(Row(ban: "запрет, доп согл"), exceptions);

        Assert.Equal(RestockSchema.NoteApprove, decision.Note);
    }

    [Fact]
    public void Исключения_КороткоеЗначениеНеЛовитВесьАцр()
    {
        var exceptions = new RestockExceptions(new[] { "369" });

        Assert.False(exceptions.Covers("ОБУВЬ", "КЕДЫ", "369-146", "369-146РОЗОВЫЙXS"));
    }

    // ===== «#н/д» =====

    [Fact]
    public void НетВСезонномФайле_ГрузимСпокойно()
    {
        Assert.Equal(RestockSchema.NoteOk, RestockBanRules.Decide(Row(ban: "#Н/Д")).Note);
    }

    [Fact]
    public void ОшибкаФормулыСчитаетсяЗаНетВФайле()
    {
        // Из COM «#Н/Д» приходит большим отрицательным числом, а не текстом.
        Assert.Equal(RestockSchema.NoteOk, RestockBanRules.Decide(Row(ban: -2146826246)).Note);
    }

    // ===== Пустой «Запрет» =====

    [Fact]
    public void ПустойЗапрет_ОстаткаХватает_Ок()
    {
        var decision = RestockBanRules.Decide(Row(quantity: 5d, available: 70d));

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.False(decision.NeedsApproval);
    }

    [Fact]
    public void ПустойЗапрет_ОстаткаНеХватает_НаСогласование()
    {
        var decision = RestockBanRules.Decide(Row(quantity: 40d, available: 30d));

        Assert.Equal(RestockSchema.NoteApprove, decision.Note);
        Assert.True(decision.NeedsApproval);
        Assert.Contains("10", decision.Reason);
    }

    [Fact]
    public void ПустойЗапрет_РовноНоль_Ок()
    {
        // Разница ноль - хватает ровно столько, сколько просят.
        var decision = RestockBanRules.Decide(Row(quantity: 20d, available: 20d));

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.False(decision.NeedsApproval);
        Assert.Null(decision.Problem);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(0d)]
    public void ПустойЗапрет_СобиратьНечего_Ок(double? available)
    {
        // Разница считается только там, где «Фактическое кол-во» заполнено: так и делает
        // аналитик, фильтруя лист по непустым значениям этой колонки.
        var decision = RestockBanRules.Decide(Row(quantity: 20d, available: available));

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.False(decision.NeedsApproval);
    }

    [Fact]
    public void ПустойЗапрет_НетКоличества_РешенияНет()
    {
        var decision = RestockBanRules.Decide(Row(quantity: null));

        Assert.Equal(string.Empty, decision.Note);
        Assert.NotNull(decision.Problem);
    }

    // ===== «доп согл» =====

    [Theory]
    [InlineData("доп согл")]
    [InlineData("запрет, доп согл")]
    [InlineData("отгрузка в рамках заказа МП, запрет забора из розницы, доп согл")]
    [InlineData("запрет забора из розницы, доп согл")]
    public void ДопСогл_ВсегдаНаСогласование(string ban)
    {
        // Проверено по готовой подтоварке: все 14 таких строк аналитик пометил
        // «Согласовать» и вынес на отдельный лист.
        var decision = RestockBanRules.Decide(Row(ban: ban, quantity: 20d, available: 100d));

        Assert.Equal(RestockSchema.NoteApprove, decision.Note);
        Assert.True(decision.NeedsApproval);
        Assert.Contains(ban, decision.Reason);
    }

    [Fact]
    public void НеизвестноеЗначениеЗапрета_РешенияНет()
    {
        var decision = RestockBanRules.Decide(Row(ban: "нельзя брать по средам"));

        Assert.Equal(string.Empty, decision.Note);
        Assert.Contains("нельзя брать по средам", decision.Problem);
    }
}
