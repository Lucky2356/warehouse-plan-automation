using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class SupplyReferenceRepairTests
{
    [Fact]
    public void ВтораяПоставкаСмотритВоВторуюКолонку() =>
        Assert.Equal(
            "='для цен'!$C$6/'для цен'!$C$4",
            SupplyReferenceRepair.Retarget("='для цен'!$B$6/'для цен'!$B$4", "C"));

    [Fact]
    public void СтрокаВернуласьКПервойПоставке() =>
        Assert.Equal(
            "='для цен'!$B$6/'для цен'!$B$4",
            SupplyReferenceRepair.Retarget("='для цен'!$C$6/'для цен'!$C$4", "B"));

    [Fact]
    public void МенятьНечего_ВозвращаетсяNull() =>
        Assert.Null(SupplyReferenceRepair.Retarget("='для цен'!$B$6/'для цен'!$B$4", "B"));

    [Fact]
    public void ЗакреплённаяСтрокаБезКолонкиНеТрогается() =>
        Assert.Null(SupplyReferenceRepair.Retarget("=СУММ(A$4:A$9)", "C"));

    [Fact]
    public void ПустаяФормула_ВозвращаетсяNull() =>
        Assert.Null(SupplyReferenceRepair.Retarget(null, "C"));
}
