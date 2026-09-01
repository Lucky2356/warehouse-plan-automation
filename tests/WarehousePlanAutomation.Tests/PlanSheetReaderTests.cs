using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Tests.TestData;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class PlanSheetReaderTests
{
    [Fact]
    public void НаходитЗаголовкиДажеСДополнительнымТекстом()
    {
        var layout = PlanFixture.BuildLayout();

        Assert.Equal(1, layout.Headers.HeaderRow);
        Assert.Equal(15, layout.Headers[SheetSchema.Plan.NetworkDate]);
        Assert.Equal(16, layout.Headers[SheetSchema.Plan.ShipmentDate]);
        Assert.Equal(5, layout.Headers[SheetSchema.Plan.Comments]);
        Assert.Equal(13, layout.Headers[SheetSchema.Plan.Status]);
        Assert.Equal(14, layout.Headers[SheetSchema.Plan.CompletionPercent]);
    }

    [Fact]
    public void РазбираетБлокиИОсобыеСтроки()
    {
        var layout = PlanFixture.BuildLayout();

        var allGroups = layout.Section(PlanSectionKind.AllGroups);
        Assert.NotNull(allGroups);
        Assert.Equal(2, allGroups!.HeaderRow);
        Assert.Equal(6, allGroups.DataRows.Count);
        Assert.Equal(3, allGroups.FirstDataRow);
        Assert.Equal(8, allGroups.LastDataRow);
        Assert.Equal("=SUM(J5:J8)", allGroups.AggregateFormula);

        Assert.Equal(3, layout.StorageAcceptanceRow!.ExcelRow);
        Assert.Equal(4, layout.AutoHubRow!.ExcelRow);

        Assert.Equal(14, layout.MarketplaceFromStorage!.ExcelRow);
        Assert.Equal(15, layout.MarketplaceFromReturns!.ExcelRow);
        Assert.Equal(16, layout.MarketplaceFromSupplies!.ExcelRow);

        Assert.Equal(17, layout.Section(PlanSectionKind.Wholesale)!.HeaderRow);
        Assert.Equal(18, layout.Section(PlanSectionKind.InternetShop)!.HeaderRow);
    }

    [Fact]
    public void СтрокаПриемкиНаХранилищеНеСчитаетсяЗаказом()
    {
        var layout = PlanFixture.BuildLayout();
        var allGroups = layout.Section(PlanSectionKind.AllGroups)!;

        Assert.False(allGroups.DataRows[0].IsOrderRow);
        Assert.False(allGroups.DataRows[1].IsOrderRow);
        Assert.True(allGroups.DataRows[2].IsOrderRow);
    }

    [Fact]
    public void ЧитаетНомераЗагрузкиИФормулы()
    {
        var layout = PlanFixture.BuildLayout();
        var urgent = layout.Section(PlanSectionKind.AllGroups)!.DataRows[2];

        Assert.Equal(PlanFixture.ExistingUrgentLoadNumber, urgent.LoadNumber);
        Assert.Equal(83d, urgent.CompletionPercent);
        Assert.Contains(SheetSchema.Plan.NetworkDate, urgent.FormulaColumns);
    }

    [Fact]
    public void ОтсутствиеОбязательнойСтрокиДаётПонятнуюОшибку()
    {
        var rows = new List<object?[]>
        {
            new object?[]
            {
                "№", "Поставки", "Обработка", "Группа", "Комментарий", "Дата документа", "Сроки выполнения",
                "Дней в работе", "Номер загрузки", "Количество единиц", "Цены", "Приоритеты", "статус",
                "% выполнения", "Дата в сети", "Дата отгрузки", "Решение",
            },
            new object?[] { "все группы" },
        };

        var grid = SheetGrid.FromRows(1, 1, rows);

        var exception = Assert.Throws<WorkbookValidationException>(() => PlanSheetReader.Read(grid));
        Assert.Contains("возвраты", exception.Message);
    }

    [Fact]
    public void ОтсутствиеКолонкиДаётПонятнуюОшибку()
    {
        var rows = new List<object?[]>
        {
            new object?[] { "№", "Поставки", "Обработка" },
        };

        var grid = SheetGrid.FromRows(1, 1, rows);

        var exception = Assert.Throws<WorkbookValidationException>(() => PlanSheetReader.Read(grid));
        Assert.Contains("Количество единиц", exception.Message);
    }
}
