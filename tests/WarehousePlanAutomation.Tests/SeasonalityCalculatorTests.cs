using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class SeasonalityCalculatorTests
{
    /// <summary>
    /// Реальные недельные доли подгруппы «МАСКА ДЛЯ СНА/» из разобранного файла:
    /// по ним и проверялась догадка о том, как считаются доли месяцев.
    /// </summary>
    private static SeasonalityRow Mask() => new(
        "МЕЛКИЕ АКСЕССУАРЫ",
        "РАЗНОЕ",
        "МАСКА ДЛЯ СНА/",
        new Dictionary<int, double>
        {
            [40] = 0.014, [41] = 0.010, [42] = 0.013, [43] = 0.011, [44] = 0.013922,
            [45] = 0.011094, [46] = 0.008484, [47] = 0.016533, [48] = 0.028497,
            [49] = 0.046987, [50] = 0.077007, [51] = 0.123994, [52] = 0.242767,
        });

    /// <summary>Те же доли плюс колонка первой недели: ею заканчивается декабрь.</summary>
    private static SeasonalityRow MaskWithFirstWeek()
    {
        var weeks = Mask().Weeks.ToDictionary(pair => pair.Key, pair => pair.Value);
        weeks[1] = 0.02d;
        return new SeasonalityRow("МЕЛКИЕ АКСЕССУАРЫ", "РАЗНОЕ", "МАСКА ДЛЯ СНА/", weeks);
    }

    private static SubgroupKey MaskKey() => new("РАЗНОЕ", "МАСКА ДЛЯ СНА/");

    [Fact]
    public void ДоляМесяцаЭтоСуммаЕгоНедель()
    {
        var plan = SeasonalityCalculator.Build(
            new[] { MaskKey() },
            new[] { new YearMonth(2026, 10), new YearMonth(2026, 11) },
            new[] { Mask() });

        var shares = plan.Rows.Single().Shares;
        Assert.Equal(0.061922d, shares[0]!.Value, 6);
        Assert.Equal(0.064608d, shares[1]!.Value, 6);
    }

    [Fact]
    public void ДекабрьЗахватываетПервуюНеделюСледующегоГода()
    {
        // Неделя 28 декабря по производственному календарю уже первая, и колонка «1»
        // на листе «КС» есть - её доля идёт в декабрь.
        var plan = SeasonalityCalculator.Build(
            new[] { MaskKey() },
            new[] { new YearMonth(2026, 12) },
            new[] { MaskWithFirstWeek() });

        Assert.Equal(0.510755d, plan.Rows.Single().Shares[0]!.Value, 6);
        Assert.Empty(plan.Warnings);
    }

    [Fact]
    public void ОтсутствующаяНеделя_ЗанижаетДолюИДаётЗамечание()
    {
        // В образце нет колонки первой недели, а декабрь её захватывает.
        var plan = SeasonalityCalculator.Build(
            new[] { MaskKey() },
            new[] { new YearMonth(2026, 12) },
            new[] { Mask() });

        Assert.Equal(0.490755d, plan.Rows.Single().Shares[0]!.Value, 6);
        Assert.Contains(plan.Warnings, w => w.Message.Contains("недели 1"));
        Assert.Contains(plan.Warnings, w => w.Message.Contains("занижена"));
    }

    [Fact]
    public void ПодгруппаНеНайдена_ДоляПустаяИЕстьЗамечание()
    {
        var plan = SeasonalityCalculator.Build(
            new[] { new SubgroupKey("РАЗНОЕ", "ЗОНТ/") },
            new[] { new YearMonth(2026, 10) },
            new[] { Mask() });

        Assert.Null(plan.Rows.Single().Shares[0]);
        Assert.Contains(plan.Warnings, w => w.Message.Contains("ЗОНТ/"));
    }

    [Fact]
    public void ПодгруппаИзДругойГруппы_НеПодставляется()
    {
        // Одна и та же подгруппа встречается в разных группах, и доли у них разные.
        var plan = SeasonalityCalculator.Build(
            new[] { new SubgroupKey("БРАСЛЕТ", "МАСКА ДЛЯ СНА/") },
            new[] { new YearMonth(2026, 10) },
            new[] { Mask() });

        Assert.Null(plan.Rows.Single().Shares[0]);
    }

    [Fact]
    public void ПовторыСОдинаковымиЧислами_БерутсяКакОдна()
    {
        var plan = SeasonalityCalculator.Build(
            new[] { MaskKey() },
            new[] { new YearMonth(2026, 10) },
            new[] { Mask(), Mask() });

        Assert.Equal(0.061922d, plan.Rows.Single().Shares[0]!.Value, 6);
        Assert.Empty(plan.Warnings);
    }

    [Fact]
    public void ПовторыСРазнымиЧислами_НеВыбираютсяАвтоматически()
    {
        var other = Mask() with { Weeks = new Dictionary<int, double> { [40] = 0.5 } };

        var plan = SeasonalityCalculator.Build(
            new[] { MaskKey() },
            new[] { new YearMonth(2026, 10) },
            new[] { Mask(), other });

        Assert.Null(plan.Rows.Single().Shares[0]);
        Assert.Contains(plan.Warnings, w => w.Message.Contains("не выбирает"));
    }
}
