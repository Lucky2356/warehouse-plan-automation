using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class TextParsingTests
{
    [Theory]
    [InlineData("2412-015 Пуховики_получение в рознице 20.09", "2412-015")]
    [InlineData("МЗ352-133 отгрузка по готовности", "352-133")]
    [InlineData("2446-001, 326-060 ШПП_отгрузка по готовности", "2446-001")]
    [InlineData("СЗ 2437-027, СЗ2433-010 (обувь)", "2437-027")]
    public void ShipmentCodeParser_НаходитНомерПоставки(string text, string expectedFirstCode)
    {
        Assert.True(ShipmentCodeParser.ContainsCode(text));
        Assert.Equal(expectedFirstCode, ShipmentCodeParser.ExtractCodes(text)[0]);
    }

    [Fact]
    public void ShipmentCodeParser_НаходитВсеНомераБезПовторов()
    {
        var codes = ShipmentCodeParser.ExtractCodes("352-132, 357-041, 326-065 ШПП, повтор 352-132");

        Assert.Equal(new[] { "352-132", "357-041", "326-065" }, codes);
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
            "Срочная подтоварка 28.08_Хранение, хранилище Номер загрузки 44575395 <Подбор:>";

        Assert.True(LoadNumberParser.TryExtract(comment, out var loadNumber));
        Assert.Equal(44575395L, loadNumber);
    }

    [Theory]
    [InlineData("Бижутерия с хранилища_срочная отгрузка НОМЕР ЗАГРУЗКИ 44507522", 44507522L)]
    [InlineData("Подтоварка номер загрузки 44366323 <Подбор:>", 44366323L)]
    public void LoadNumberParser_НеЗависитОтРегистра(string comment, long expected)
    {
        Assert.True(LoadNumberParser.TryExtract(comment, out var loadNumber));
        Assert.Equal(expected, loadNumber);
    }

    [Theory]
    [InlineData("на образцы, хранение на складе СЗ369-062 (тапки)")]
    [InlineData("виртуальный возврат")]
    public void LoadNumberParser_НеНаходитНомерЕслиЕгоНет(string comment)
    {
        Assert.False(LoadNumberParser.TryExtract(comment, out _));
    }

    [Theory]
    [InlineData("Подтоварка Номер загрузки не указан, отгрузка 05.09")]
    [InlineData("Номер загрузки уточняется 44575395")]
    public void LoadNumberParser_НеБерётДалёкоеЧислоЕслиНомераПослеМаркераНет(string comment)
    {
        // Между словами «Номер загрузки» и номером допускаются только разделители:
        // иначе номером загрузки стало бы первое попавшееся дальше число.
        Assert.False(LoadNumberParser.TryExtract(comment, out _));
    }

    [Theory]
    [InlineData("Подтоварка Номер загрузки: 44575395", 44575395L)]
    [InlineData("Подтоварка Номер загрузки № 44575395", 44575395L)]
    public void LoadNumberParser_ДопускаетРазделителиПередНомером(string comment, long expected)
    {
        Assert.True(LoadNumberParser.TryExtract(comment, out var loadNumber));
        Assert.Equal(expected, loadNumber);
    }

    [Fact]
    public void LoadNumberParser_ВозвращаетТекстПоставокДоСловНомерЗагрузки()
    {
        const string comment =
            "Срочная подтоварка 28.08_Хранение, хранилище Номер загрузки 44575395 <Подбор:>";

        Assert.Equal(
            "Срочная подтоварка 28.08_Хранение, хранилище",
            LoadNumberParser.ExtractSuppliesText(comment));
    }

    [Theory]
    [InlineData("ЗП373-1739")]
    [InlineData("Отгрузка ЗП373-1739 Номер загрузки 44388195")]
    [InlineData("зп 373-1739")]
    public void OrderTextRules_НаходитДокументЗП(string comment)
    {
        Assert.True(OrderTextRules.IsZpRow(comment));
    }

    [Theory]
    [InlineData("2412-015 Пуховики_получение в рознице 20.09")]
    [InlineData("Срочная подтоварка_Хранение, хранилище")]
    public void OrderTextRules_НеПутаетЗПСЧастямиДругихСлов(string comment)
    {
        // «ЗП» ищется как отдельный признак: две буквы слишком коротки,
        // чтобы искать их подстрокой.
        Assert.False(OrderTextRules.IsZpRow(comment));
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
        var text = "Номер" + noBreakSpace + "загрузки 44366323";

        Assert.Equal("номер загрузки 44366323", TextUtils.NormalizeKey(text));
        Assert.True(LoadNumberParser.TryExtract(text, out var loadNumber));
        Assert.Equal(44366323L, loadNumber);
    }

    [Fact]
    public void TextUtils_ПриводитДлинныеЧислаБезЭкспоненты()
    {
        Assert.Equal("44575395", TextUtils.CellToString(44575395d));
    }
}
