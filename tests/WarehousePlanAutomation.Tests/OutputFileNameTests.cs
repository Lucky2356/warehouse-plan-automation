using WarehousePlanAutomation.Core.Models;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class OutputFileNameTests
{
    private static readonly DateTime Now = new(2026, 9, 9, 15, 50, 22);

    [Fact]
    public void КИмениДобавляетсяЧтоСделалиИКогда() =>
        Assert.Equal(
            "C2518-084 Бижутерия LTL719 подготовка 2026-09-09_1550",
            OutputFileName.Build("C2518-084 Бижутерия LTL719", OutputFileName.PrepareMark, Now));

    [Fact]
    public void ОтметкаПрошлогоЭтапаНеНакапливается()
    {
        // Файл второго этапа делается из файла первого.
        var prepared = OutputFileName.Build("Поставка", OutputFileName.PrepareMark, Now);

        Assert.Equal(
            "Поставка распред 2026-09-09_1552",
            OutputFileName.Build(prepared, OutputFileName.DistributionMark, Now.AddMinutes(2)));
    }

    [Fact]
    public void СтароеИмяСГотовоТожеПереименовываетсяНачисто() =>
        Assert.Equal(
            "Primer_raspred подготовка 2026-09-09_1550",
            OutputFileName.Build("Primer_raspred_готово_2026-09-08_012823", OutputFileName.PrepareMark, Now));

    [Fact]
    public void ОтметкаСНомеромПовтораТожеСнимается() =>
        Assert.Equal(
            "Поставка",
            OutputFileName.WithoutMark("Поставка распред 2026-09-09_1552_2"));

    [Fact]
    public void ЧужойХвостНеТрогается()
    {
        // Дата в названии поставки - это не отметка программы.
        Assert.Equal(
            "Заказы на отгрузку 04_09",
            OutputFileName.WithoutMark("Заказы на отгрузку 04_09"));

        Assert.Equal(
            "Отчёт 2026-09-09",
            OutputFileName.WithoutMark("Отчёт 2026-09-09"));
    }

    [Fact]
    public void ПланСкладаПодписываетсяСвоимСловом() =>
        Assert.Equal(
            "Plan_sklada_03_09 план склада 2026-09-09_1550",
            OutputFileName.Build("Plan_sklada_03_09", OutputFileName.PlanMark, Now));
}
