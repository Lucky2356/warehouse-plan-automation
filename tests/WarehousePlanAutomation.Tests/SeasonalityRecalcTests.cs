using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class SeasonalityRecalcTests
{
    [Theory]
    [InlineData("БИЖУТЕРИЯ", 2)]
    [InlineData("бижутерия", 2)]
    [InlineData("КОЛГОТКИ,НОСКИ", 3)]
    [InlineData("КОЛГОТКИ, НОСКИ", 3)]
    [InlineData("УКРАШЕНИЯ ДЛЯ ВОЛОС", 3)]
    [InlineData("МЕЛКИЕ АКСЕССУАРЫ", 0)]
    [InlineData("", 0)]
    [InlineData(null, 0)]
    public void МесяцыПересчётаЗависятОтСектора(string? sector, int expected) =>
        Assert.Equal(expected, SeasonalityRecalc.MonthsFor(sector));

    [Fact]
    public void ДваМесяца_ТретийОстаётсяКакБыл()
    {
        // Шаблон «Бижу»: 4 % и 6 % превращаются в 40 % и 60 %.
        var shares = SeasonalityRecalc.Apply(new double?[] { 0.04d, 0.06d, 0.07d }, 2);

        Assert.Equal(0.4d, shares[0]!.Value, 9);
        Assert.Equal(0.6d, shares[1]!.Value, 9);
        Assert.Equal(0.07d, shares[2]!.Value, 9);
    }

    [Fact]
    public void ТриМесяца_СуммаСтановитсяЕдиницей()
    {
        // Шаблон «Носки, укр»: 13 %, 30 % и 22 % - это 20 %, 46,15 % и 33,85 %.
        var shares = SeasonalityRecalc.Apply(new double?[] { 0.13d, 0.30d, 0.22d }, 3);

        Assert.Equal(0.2d, shares[0]!.Value, 4);
        Assert.Equal(0.4615d, shares[1]!.Value, 4);
        Assert.Equal(0.3385d, shares[2]!.Value, 4);
        Assert.Equal(1d, shares.Sum(share => share ?? 0d), 9);
    }

    [Fact]
    public void ПустаяДоляСчитаетсяНулём()
    {
        var shares = SeasonalityRecalc.Apply(new double?[] { 0d, 0.13d, 0.32d }, 3);

        Assert.Equal(0d, shares[0]!.Value, 9);
        Assert.Equal(0.13d / 0.45d, shares[1]!.Value, 9);
    }

    [Fact]
    public void СуммаНоль_ДолиНеТрогаются()
    {
        var source = new double?[] { 0d, null, 0.07d };

        Assert.Equal(source, SeasonalityRecalc.Apply(source, 2));
    }

    [Fact]
    public void СекторБезПересчёта_ДолиНеТрогаются()
    {
        var source = new double?[] { 0.04d, 0.06d, 0.07d };

        Assert.Equal(source, SeasonalityRecalc.Apply(source, 0));
    }
}
