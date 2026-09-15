using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class StockReferenceRepairTests
{
    [Fact]
    public void ДиапазонПоКолонкам_ПолучаетНовуюПоследнююКолонку()
    {
        // Строка одна и та же слева и справа - меняется только колонка.
        Assert.Equal(
            "=SUMPRODUCT((остатки!$C$14:$AZ$14)*(остатки!$C$1:$AZ$1=CONCATENATE(N$3&N$4&N$19)))",
            StockReferenceRepair.Retarget(
                "=SUMPRODUCT((остатки!$C$14:$NZ$14)*(остатки!$C$1:$NZ$1=CONCATENATE(N$3&N$4&N$19)))",
                "остатки",
                393,
                ExcelColumn.FromLetters("AZ")));
    }

    [Fact]
    public void ДиапазонПоСтрокамИКолонкам_ПолучаетОбеГраницы()
    {
        Assert.Equal(
            "=SUMPRODUCT((остатки!$C$35:$AZ$400)*(остатки!$A$35:$A$400=TEXT($A28,\"000\")))",
            StockReferenceRepair.Retarget(
                "=SUMPRODUCT((остатки!$C$35:$NZ$393)*(остатки!$A$35:$A$393=TEXT($A28,\"000\")))",
                "остатки",
                400,
                ExcelColumn.FromLetters("AZ")));
    }

    [Fact]
    public void ДиапазонВОднуКолонку_КолонкуНеМеняет()
    {
        // «остатки!$A$35:$A$393» - это колонка кодов магазинов, ей ширина не нужна.
        Assert.Equal(
            "=SUM(остатки!$A$35:$A$400)",
            StockReferenceRepair.Retarget("=SUM(остатки!$A$35:$A$393)", "остатки", 400, 50));
    }

    [Fact]
    public void ЧужойЛистНеТрогается()
    {
        Assert.Null(StockReferenceRepair.Retarget(
            "=SUMPRODUCT((Цены!$C$14:$NZ$14))", "остатки", 400, 50));
    }

    [Fact]
    public void ИмяЛистаВКавычках_ТожеУзнаётся()
    {
        Assert.Equal(
            "='для цен'!$C$1:$AZ$1",
            StockReferenceRepair.Retarget("='для цен'!$C$1:$NZ$1", "для цен", 10, ExcelColumn.FromLetters("AZ")));
    }

    [Fact]
    public void ГраницыУжеВерные_НичегоНеМеняется()
    {
        Assert.Null(StockReferenceRepair.Retarget(
            "=SUM(остатки!$C$1:$NZ$1)", "остатки", 393, ExcelColumn.FromLetters("NZ")));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("остатки")]
    public void НеФормулаНеТрогается(string? formula)
    {
        Assert.Null(StockReferenceRepair.Retarget(formula, "остатки", 400, 50));
    }
}
