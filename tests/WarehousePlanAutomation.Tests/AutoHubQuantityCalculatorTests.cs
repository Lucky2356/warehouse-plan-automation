using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class AutoHubQuantityCalculatorTests
{
    [Theory]
    [InlineData(DayOfWeek.Monday, 24000d)]
    [InlineData(DayOfWeek.Tuesday, 18000d)]
    [InlineData(DayOfWeek.Wednesday, 12000d)]
    [InlineData(DayOfWeek.Thursday, 6000d)]
    [InlineData(DayOfWeek.Friday, 42000d)]
    public void РабочиеДни_ДаютФиксированноеЗначение(DayOfWeek day, double expected)
    {
        Assert.Equal(expected, AutoHubQuantityCalculator.GetQuantity(day));
    }

    [Theory]
    [InlineData(DayOfWeek.Saturday)]
    [InlineData(DayOfWeek.Sunday)]
    public void Выходные_ЗначениеНеМеняется(DayOfWeek day)
    {
        Assert.Null(AutoHubQuantityCalculator.GetQuantity(day));
    }
}
