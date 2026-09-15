using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class PreflightCheckTests
{
    private static WorkbookProbe Probe(
        IEnumerable<string>? sheets = null,
        IEnumerable<string>? links = null) =>
        new(
            (sheets ?? new[] { "Заказы на отгрузку 04_09", "Журнал заказов на отгрузку", "План" }).ToList(),
            (links ?? Array.Empty<string>()).ToList());

    private static readonly string[] PlanSheets = { "План", "Заказы на отгрузку", "Журнал заказов на отгрузку" };

    [Fact]
    public void ВсёНаМесте_ЗамечанийНет()
    {
        Assert.Empty(PreflightCheck.Run(Probe(), PlanSheets, _ => true));
    }

    [Fact]
    public void НетНеобязательногоЛиста_ТолькоПредупреждение()
    {
        var optional = new Dictionary<string, string>
        {
            ["Р"] = "места хранения и загрузочники не соберутся",
            ["МПП"] = "места хранения и загрузочники не соберутся",
        };

        var issues = PreflightCheck.Run(
            Probe(sheets: new[] { "на загрузку", "МПП", "Согласовать" }), new[] { "на загрузку" }, _ => true, optional);

        var issue = Assert.Single(issues);
        Assert.False(issue.IsProblem);
        Assert.Contains("«Р»", issue.Message);
    }

    [Fact]
    public void НетЛиста_ЭтоПомеха()
    {
        var issues = PreflightCheck.Run(
            Probe(sheets: new[] { "План", "Заказы на отгрузку" }), PlanSheets, _ => true);

        var issue = Assert.Single(issues);
        Assert.True(issue.IsProblem);
        Assert.Contains("Журнал заказов на отгрузку", issue.Message);
    }

    [Fact]
    public void ДваПодходящихЛиста_ЭтоПомехаИОбаНазваны()
    {
        var issues = PreflightCheck.Run(
            Probe(sheets: new[]
            {
                "План", "Журнал заказов на отгрузку",
                "Заказы на отгрузку 03_09", "Заказы на отгрузку 04_09",
            }),
            PlanSheets,
            _ => true);

        var issue = Assert.Single(issues);
        Assert.True(issue.IsProblem);
        Assert.Contains("Заказы на отгрузку 03_09", issue.Message);
        Assert.Contains("Заказы на отгрузку 04_09", issue.Message);
    }

    [Fact]
    public void НесколькоИнвойсов_НеПомеха()
    {
        // Инвойс приходит листом на каждую поставку, обработка берёт все.
        var sheets = new[] { "Invoice-338", "Invoice-341", "для цен", "Цены" };

        Assert.Empty(PreflightCheck.Run(
            Probe(sheets: sheets), new[] { "Invoice", "для цен", "Цены" }, _ => true,
            repeatableSheets: new[] { "Invoice" }));

        var issue = Assert.Single(PreflightCheck.Run(
            Probe(sheets: sheets.Skip(2)), new[] { "Invoice", "для цен", "Цены" }, _ => true,
            repeatableSheets: new[] { "Invoice" }));
        Assert.True(issue.IsProblem);
    }

    [Fact]
    public void ТочноеСовпадениеСильнееНачалаНазвания()
    {
        // «прайс» и «Прайс по подразделениям» начинаются одинаково: без этого правила
        // они вечно считались бы двумя кандидатами на одно имя.
        var issues = PreflightCheck.Run(
            Probe(sheets: new[] { "прайс", "Прайс по подразделениям" }),
            new[] { "прайс", "Прайс по подраздел" },
            _ => true);

        Assert.Empty(issues);
    }

    [Fact]
    public void СвязанныйФайлНедоступен_ЭтоПредупреждение()
    {
        // Не помеха: обработка пройдёт, просто колонки со стенками не пересчитаются.
        var issues = PreflightCheck.Run(
            Probe(links: new[] { @"\\FileServer\склад$\FW26-27\стенки.xlsx" }),
            PlanSheets,
            _ => false);

        var issue = Assert.Single(issues);
        Assert.False(issue.IsProblem);
        Assert.Contains(@"\\FileServer\склад$\FW26-27\стенки.xlsx", issue.Message);
    }

    [Fact]
    public void СвязанныйФайлНаМесте_НичегоНеПишет()
    {
        Assert.Empty(PreflightCheck.Run(
            Probe(links: new[] { @"\\FileServer\склад$\FW26-27\стенки.xlsx" }), PlanSheets, _ => true));
    }
}
