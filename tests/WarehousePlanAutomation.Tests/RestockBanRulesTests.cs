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
        double? sellout = 0d,
        string sector = "ОБУВЬ") =>
        new(0, sector, "369-146РОЗОВЫЙXS", quantity, ban, available, sellout);

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
    public void ЗапретРозницы_ОбычныйСектор_НаСогласование()
    {
        var decision = RestockBanRules.Decide(Row(ban: NoRetail, sector: "ОБУВЬ"));

        Assert.Equal(RestockSchema.NoteApprove, decision.Note);
        Assert.True(decision.Highlight);
        Assert.True(decision.NeedsApproval);
    }

    [Theory]
    [InlineData("БИЖУТЕРИЯ")]
    [InlineData("МЕЛКИЕ АКСЕССУАРЫ")]
    [InlineData("мелкие аксессуары")]
    [InlineData("УКРАШЕНИЯ ДЛЯ ВОЛОС")]
    public void ЗапретРозницы_МелочьНеСогласовывается(string sector)
    {
        var decision = RestockBanRules.Decide(Row(ban: NoRetail, sector: sector));

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.True(decision.Highlight);
        Assert.False(decision.NeedsApproval);
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
    public void ПустойЗапрет_РовноНоль_ВысокийSellout_НаСогласование()
    {
        var decision = RestockBanRules.Decide(Row(quantity: 20d, available: 20d, sellout: 80d));

        Assert.Equal(RestockSchema.NoteApprove, decision.Note);
        Assert.True(decision.NeedsApproval);
        Assert.False(decision.Highlight);
    }

    [Fact]
    public void ПустойЗапрет_РовноНоль_НизкийSellout_Ок()
    {
        var decision = RestockBanRules.Decide(Row(quantity: 20d, available: 20d, sellout: 79.9d));

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.False(decision.NeedsApproval);
    }

    [Fact]
    public void ПустойЗапрет_РовноНоль_ВысокийSellout_НоМелочь_Ок()
    {
        var decision = RestockBanRules.Decide(
            Row(quantity: 20d, available: 20d, sellout: 100d, sector: "БИЖУТЕРИЯ"));

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.False(decision.NeedsApproval);
    }

    [Fact]
    public void ПустойЗапрет_РовноНоль_БезSellout_ОставляетОкИДаётЗамечание()
    {
        // Утверждать, что sellout высокий, не из чего: строка остаётся «Ок»,
        // но замечание о непрочитанном sellout пишется.
        var decision = RestockBanRules.Decide(Row(quantity: 20d, available: 20d, sellout: null));

        Assert.Equal(RestockSchema.NoteOk, decision.Note);
        Assert.False(decision.NeedsApproval);
        Assert.NotNull(decision.Problem);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(0d)]
    public void ПустойЗапрет_СобиратьНечего_РешаетSellout(double? available)
    {
        // Разница считается только там, где «Фактическое кол-во» заполнено:
        // так и делает аналитик, фильтруя лист по непустым значениям этой колонки.
        var low = RestockBanRules.Decide(Row(quantity: 20d, available: available, sellout: 5d));
        var high = RestockBanRules.Decide(Row(quantity: 20d, available: available, sellout: 95d));
        var small = RestockBanRules.Decide(
            Row(quantity: 20d, available: available, sellout: 95d, sector: "БИЖУТЕРИЯ"));

        Assert.Equal(RestockSchema.NoteOk, low.Note);
        Assert.False(low.NeedsApproval);

        Assert.Equal(RestockSchema.NoteApprove, high.Note);
        Assert.Contains("собрать на ВБ+Озон нечего", high.Reason);

        Assert.Equal(RestockSchema.NoteOk, small.Note);
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
