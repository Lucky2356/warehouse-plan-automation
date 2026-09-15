using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class TotalsFormulaRepairTests
{
    // Данные в строках 2..15, итог в 16 - так выглядит «Цены» после добавления строк.
    private const int First = 2;
    private const int Last = 15;
    private const int Totals = 16;

    [Theory]
    [InlineData("=SUBTOTAL(9,P2:P2)", "=SUBTOTAL(9,P2:P15)")]
    [InlineData("=SUM(P2:P3)", "=SUM(P2:P15)")]
    [InlineData("=AVERAGE($Q$2:$Q$3)", "=AVERAGE($Q$2:$Q$15)")]
    [InlineData("=COUNTIF(B2:B3,\"431-329\")", "=COUNTIF(B2:B15,\"431-329\")")]
    [InlineData("=SUMPRODUCT(P1:P3,R1:R3)", "=SUMPRODUCT(P1:P15,R1:R15)")]
    public void ФормулаПодТаблицей_РастягиваетсяНаВсеСтроки(string formula, string expected)
    {
        Assert.Equal(expected, TotalsFormulaRepair.ExtendBelow(formula, First, Last, Totals));
    }

    [Theory]
    [InlineData("=SUM(P2:P15)")]              // уже верно
    [InlineData("=AG16/AF16")]                // ссылки на итоговую строку
    [InlineData("=SUM(P2:P16)")]              // захватывает итог - это не диапазон данных
    [InlineData("=SUM(P5:P8)")]               // начинается не с первой строки
    [InlineData("='для цен'!B5:B7")]          // другой лист
    [InlineData("=SUM(Остатки!A2:A3)")]
    [InlineData("2665")]
    [InlineData(null)]
    public void ОстальноеНеТрогается(string? formula)
    {
        Assert.Null(TotalsFormulaRepair.ExtendBelow(formula, First, Last, Totals));
    }
}

public class UrgentRestockFormulaTests
{
    [Theory]
    [InlineData("Срочная подтоварка 28.08_Хранение, хранилище", true)]
    [InlineData("СРОЧНАЯ ПОДТОВАРКА ОБУВИ", true)]
    [InlineData("Бижутерия с хранилища_срочная отгрузка", false)]
    [InlineData("Подтоварка МП", false)]
    public void СрочнаяПодтоваркаУзнаётсяПоТексту(string text, bool expected)
    {
        Assert.Equal(expected, OrderTextRules.IsUrgentRestock(text));
    }

    [Theory]
    [InlineData("=RC[-9]+R2C15", true, "=RC[-9]+R2C15-4")]
    [InlineData("=RC[-9]+R2C15-4", true, "=RC[-9]+R2C15-4")]
    [InlineData("=RC[-9]+R2C15-4", false, "=RC[-9]+R2C15")]
    [InlineData("=RC[-9]+R2C15", false, "=RC[-9]+R2C15")]
    public void ВычитаниеСтавитсяИУбираетсяПоСтроке(string donor, bool urgent, string expected)
    {
        Assert.Equal(expected, OrderTextRules.NetworkDateFormula(donor, urgent));
    }
}
