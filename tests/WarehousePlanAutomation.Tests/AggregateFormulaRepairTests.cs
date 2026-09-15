using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class AggregateFormulaRepairTests
{
    [Fact]
    public void СохраняетОтступНачалаСуммыОтНачалаБлока()
    {
        // Вчера: блок начинался со строки 3, сумма - со строки 5 (две особые строки не суммируются).
        var repaired = AggregateFormulaRepair.BuildRepairedFormula("=SUM(J5:J24)", 3, 3, 27);

        Assert.Equal("=SUM(J5:J27)", repaired);
    }

    [Fact]
    public void УчитываетСмещениеБлокаВверх()
    {
        var repaired = AggregateFormulaRepair.BuildRepairedFormula("=SUM(J5:J24)", 3, 2, 20);

        Assert.Equal("=SUM(J4:J20)", repaired);
    }

    [Fact]
    public void БлокБезОтступаРастягиваетсяНаВсеСтроки()
    {
        var repaired = AggregateFormulaRepair.BuildRepairedFormula("=SUM(J26:J28)", 26, 26, 31);

        Assert.Equal("=SUM(J26:J31)", repaired);
    }

    [Theory]
    [InlineData("=SUM(J5:J24)+100")]
    [InlineData("=SUBTOTAL(9;J5:J24)")]
    [InlineData("=SUM(J5:K24)")]
    [InlineData(null)]
    public void СложныеФормулыНеТрогаются(string? formula)
    {
        Assert.Null(AggregateFormulaRepair.BuildRepairedFormula(formula, 3, 3, 27));
    }
}
