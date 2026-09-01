using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class TextParsingTests
{
    [Theory]
    [InlineData("1022-015 Пуховики_получение в рознице 20.09", "1022-015")]
    [InlineData("МЗ806-133 отгрузка по готовности", "806-133")]
    [InlineData("1244-001, 139-060 ШПП_отгрузка по готовности", "1244-001")]
    [InlineData("СЗ 1117-027, СЗ1103-010 (обувь)", "1117-027")]
    public void ShipmentCodeParser_НаходитНомерПоставки(string text, string expectedFirstCode)
    {
        Assert.True(ShipmentCodeParser.ContainsCode(text));
        Assert.Equal(expectedFirstCode, ShipmentCodeParser.ExtractCodes(text)[0]);
    }

    [Fact]
    public void ShipmentCodeParser_НаходитВсеНомераБезПовторов()
    {
        var codes = ShipmentCodeParser.ExtractCodes("806-132, 799-041, 139-065 ШПП, повтор 806-132");

        Assert.Equal(new[] { "806-132", "799-041", "139-065" }, codes);
    }

    [Theory]
    [InlineData("Сезонный товар FW26-27 из возвратов, времянки")]
    [InlineData("Приоритет к 10.08, микс А1,А2,А3")]
    [InlineData("Заказ интерент магазина № 0124589289-0331-1")]
    [InlineData("")]
    public void ShipmentCodeParser_НеПутаетСлужебныеФрагментыСНомеромПоставки(string text)
    {
        Assert.False(ShipmentCodeParser.ContainsCode(text));
    }

    [Fact]
    public void LoadNumberParser_ИзвлекаетНомерЗагрузки()
    {
        const string comment =
            "Срочная подтоварка 28.08_Хранение, хранилище Номер загрузки 55575395 <Подбор:>";

        Assert.True(LoadNumberParser.TryExtract(comment, out var loadNumber));
        Assert.Equal(55575395L, loadNumber);
    }

    [Theory]
    [InlineData("Бижутерия с хранилища_срочная отгрузка НОМЕР ЗАГРУЗКИ 55507522", 55507522L)]
    [InlineData("Подтоварка номер загрузки 55366323 <Подбор:>", 55366323L)]
    public void LoadNumberParser_НеЗависитОтРегистра(string comment, long expected)
    {
        Assert.True(LoadNumberParser.TryExtract(comment, out var loadNumber));
        Assert.Equal(expected, loadNumber);
    }

    [Theory]
    [InlineData("на образцы, хранение на складе СЗ985-062 (тапки)")]
    [InlineData("виртуальный возврат")]
    public void LoadNumberParser_НеНаходитНомерЕслиЕгоНет(string comment)
    {
        Assert.False(LoadNumberParser.TryExtract(comment, out _));
    }

    [Theory]
    [InlineData("Подтоварка Номер загрузки не указан, отгрузка 05.09")]
    [InlineData("Номер загрузки уточняется 55575395")]
    public void LoadNumberParser_НеБерётДалёкоеЧислоЕслиНомераПослеМаркераНет(string comment)
    {
        // Между словами «Номер загрузки» и номером допускаются только разделители:
        // иначе номером загрузки стало бы первое попавшееся дальше число.
        Assert.False(LoadNumberParser.TryExtract(comment, out _));
    }

    [Theory]
    [InlineData("Подтоварка Номер загрузки: 55575395", 55575395L)]
    [InlineData("Подтоварка Номер загрузки № 55575395", 55575395L)]
    public void LoadNumberParser_ДопускаетРазделителиПередНомером(string comment, long expected)
    {
        Assert.True(LoadNumberParser.TryExtract(comment, out var loadNumber));
        Assert.Equal(expected, loadNumber);
    }

    [Fact]
    public void LoadNumberParser_ВозвращаетТекстПоставокДоСловНомерЗагрузки()
    {
        const string comment =
            "Срочная подтоварка 28.08_Хранение, хранилище Номер загрузки 55575395 <Подбор:>";

        Assert.Equal(
            "Срочная подтоварка 28.08_Хранение, хранилище",
            LoadNumberParser.ExtractSuppliesText(comment));
    }

    [Theory]
    [InlineData("Сроч подтоварка")]
    [InlineData("Срочно, обувь")]
    [InlineData("Срочная подтоварка 28.08")]
    [InlineData("СРОЧНАЯ отгрузка ШПП")]
    public void OrderTextRules_ОпределяетСрочность(string text)
    {
        Assert.True(OrderTextRules.IsUrgent(text));
    }

    [Fact]
    public void OrderTextRules_НеСчитаетСрочнымиОбычныеЗаказы()
    {
        Assert.False(OrderTextRules.IsUrgent("1022-015 Пуховики_получение в рознице 20.09"));
    }

    [Theory]
    [InlineData("1100-026, 028 Обувь СЕТ1_в рознице", SetMonoKind.Set)]
    [InlineData("1100-026, 028 Обувь СЕТ2_в рознице", SetMonoKind.Set)]
    [InlineData("1117-029 Кеды на Юг МОНО_в рознице", SetMonoKind.Mono)]
    [InlineData("1022-015 Пуховики", SetMonoKind.Neutral)]
    public void OrderTextRules_ОпределяетСетИМоно(string supplies, SetMonoKind expected)
    {
        Assert.Equal(expected, OrderTextRules.DetectSetMono(supplies));
    }

    [Theory]
    [InlineData("Обувь_дата в сети с 01.10")]
    [InlineData("Кассета для украшений")]
    [InlineData("Монолитная упаковка")]
    public void OrderTextRules_НеПутаетСетИМоноСЧастямиДругихСлов(string supplies)
    {
        // «СЕТ» и «МОНО» ищутся как отдельные признаки, а не как подстроки:
        // иначе «в сети» и «кассета» стали бы признаком СЕТ.
        Assert.Equal(SetMonoKind.Neutral, OrderTextRules.DetectSetMono(supplies));
    }

    [Fact]
    public void TextUtils_НормализуетПробелыИРегистр()
    {
        Assert.Equal("разница ед", TextUtils.NormalizeKey("  Разница  ед  "));
        Assert.Equal("приемка на хранилище", TextUtils.NormalizeKey("Приёмка   на хранилище"));
    }

    [Fact]
    public void TextUtils_ЗаменяетНеразрывныеПробелы()
    {
        var noBreakSpace = ((char)0x00A0).ToString();
        var text = "Номер" + noBreakSpace + "загрузки 55366323";

        Assert.Equal("номер загрузки 55366323", TextUtils.NormalizeKey(text));
        Assert.True(LoadNumberParser.TryExtract(text, out var loadNumber));
        Assert.Equal(55366323L, loadNumber);
    }

    [Fact]
    public void TextUtils_ПриводитДлинныеЧислаБезЭкспоненты()
    {
        Assert.Equal("55575395", TextUtils.CellToString(55575395d));
    }
}
