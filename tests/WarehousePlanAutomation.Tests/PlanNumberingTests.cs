using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Tests.TestData;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class PlanNumberingTests
{
    [Fact]
    public void НумерацияИдётПоПорядкуСтрокБлока()
    {
        var layout = PlanFixture.BuildLayout();
        var allGroups = layout.Section(PlanSectionKind.AllGroups)!;

        var numbers = PlanNumberingBuilder.Build(layout)
            .Where(n => allGroups.DataRows.Any(r => r.ExcelRow == n.ExcelRow))
            .ToList();

        Assert.Equal(allGroups.DataRows.Select(r => r.ExcelRow), numbers.Select(n => n.ExcelRow));
        Assert.Equal(new[] { 1, 2, 3, 4, 5, 6 }, numbers.Select(n => n.Number));
    }

    [Fact]
    public void КаждыйБлокНумеруетсяСЕдиницы()
    {
        var layout = PlanFixture.BuildLayout();
        var numbers = PlanNumberingBuilder.Build(layout).ToDictionary(n => n.ExcelRow, n => n.Number);

        foreach (var section in layout.OrderSections)
        {
            Assert.Equal(1, numbers[section.DataRows[0].ExcelRow]);
        }
    }

    [Fact]
    public void ПустаяСтрокаВнутриБлокаНумерациюНеЛомает()
    {
        // Пустая строка не считается строкой данных, поэтому номера остаются сплошными,
        // а сама строка номера не получает.
        var layout = PlanFixture.BuildLayoutWithGapInAllGroups();
        var allGroups = layout.Section(PlanSectionKind.AllGroups)!;

        var numbers = PlanNumberingBuilder.Build(layout)
            .Where(n => allGroups.DataRows.Any(r => r.ExcelRow == n.ExcelRow))
            .ToList();

        Assert.Equal(allGroups.DataRows.Select(r => r.ExcelRow), numbers.Select(n => n.ExcelRow));
        Assert.Equal(new[] { 1, 2, 3, 4 }, numbers.Select(n => n.Number));
    }

    [Fact]
    public void АгрегатныеБлокиНеНумеруются()
    {
        var layout = PlanFixture.BuildLayout();
        var marketplaces = layout.Section(PlanSectionKind.Marketplaces)!;
        var numbered = PlanNumberingBuilder.Build(layout).Select(n => n.ExcelRow).ToHashSet();

        Assert.DoesNotContain(marketplaces.DataRows.Select(r => r.ExcelRow), numbered.Contains);
    }
}
