using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class SectorSelectionTests
{
    private static PriceRowValues Row(int index, string? sector) => new(
        index,
        0,
        "186828020" + index,
        "431-329",
        10d,
        1d,
        10d,
        1d,
        sector is null
            ? null
            : new PriceListRow("1", "155588", "модель", sector, "РАЗНОЕ", "имя", "арт", "цвет", "сезон", "TU"),
        "подгруппа",
        null,
        false);

    [Fact]
    public void СекторыИдутПоЧислуСтрок()
    {
        var sectors = SectorSelection.Sectors(new[]
        {
            Row(0, "УКРАШЕНИЯ ДЛЯ ВОЛОС"),
            Row(1, "БИЖУТЕРИЯ"),
            Row(2, "БИЖУТЕРИЯ"),
            Row(3, null),
        });

        Assert.Equal(new[] { "БИЖУТЕРИЯ", "УКРАШЕНИЯ ДЛЯ ВОЛОС" }, sectors);
    }

    [Fact]
    public void СекторУзнаётсяВНазванииФайла()
    {
        var matched = SectorSelection.MatchFileName(
            new[] { "БИЖУТЕРИЯ", "УКРАШЕНИЯ ДЛЯ ВОЛОС" },
            @"\\FileServer\склад$\подготовка поставок\C2518-084 Бижутерия LTL719.xlsx");

        Assert.Equal("БИЖУТЕРИЯ", Assert.Single(matched));
    }

    [Fact]
    public void ХватаетОдногоСловаИзНазванияСектора()
    {
        // Сектор называется «КОЛГОТКИ,НОСКИ», а файл - просто «Носки».
        var matched = SectorSelection.MatchFileName(
            new[] { "КОЛГОТКИ,НОСКИ", "СУМКИ" }, "318-140-036 Носки FTL 36.xlsx");

        Assert.Equal("КОЛГОТКИ,НОСКИ", Assert.Single(matched));
    }

    [Fact]
    public void ДваСектораВНазвании_ОбаВозвращаются()
    {
        // Тогда по названию не решить, и программа спросит.
        var matched = SectorSelection.MatchFileName(
            new[] { "КОЛГОТКИ,НОСКИ", "СУМКИ" }, "318-135 Сумки, 327-036 Носки FTL 36.xlsx");

        Assert.Equal(2, matched.Count);
    }

    [Fact]
    public void КороткиеСловаНеСчитаются()
    {
        // «ДЛЯ» есть в половине названий и совпало бы с чем угодно.
        Assert.Empty(SectorSelection.MatchFileName(
            new[] { "УКРАШЕНИЯ ДЛЯ ВОЛОС" }, "431-329 Маски для сна.xlsx"));
    }

    [Fact]
    public void ЧастьСловаНеСчитается()
    {
        Assert.Empty(SectorSelection.MatchFileName(new[] { "ОЧКИ" }, "431-329 Очкииии.xlsx"));
    }

    [Fact]
    public void БезНазванияФайлаНичегоНеНаходится()
    {
        Assert.Empty(SectorSelection.MatchFileName(new[] { "БИЖУТЕРИЯ" }, null));
    }
}
