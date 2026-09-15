using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class LoaderFormulaRepairTests
{
    // Колонки кодов после расширения - G..I (7..9).
    private const int First = 7;
    private const int Last = 9;

    [Theory]
    [InlineData("=SUM(G13:G13)")]
    [InlineData("=SUM(H13:H13)")]
    [InlineData("=SUM(I13:I13)")]
    public void СуммаИзОднойЯчейкиРастягиваетсяНаВсеКолонки(string formula)
    {
        // Так выглядит сумма по колонкам кодов, когда колонка была одна: Excel
        // диапазон из одной ячейки не расширяет, куда бы ни вставили новые колонки.
        Assert.Equal("=SUM(G13:I13)", LoaderFormulaRepair.Expand(formula, First, Last));
    }

    [Fact]
    public void СуммаПоКолонкеНеТрогается()
    {
        // «СУММ(G13:G131)» считает свою колонку сверху вниз - её растягивать нельзя.
        Assert.Null(LoaderFormulaRepair.Expand("=SUM(I13:I131)", First, Last));
    }

    [Fact]
    public void УжеРастянутаяСуммаНеТрогается()
    {
        Assert.Null(LoaderFormulaRepair.Expand("=SUM(G13:I13)", First, Last));
    }

    [Fact]
    public void ДиапазонВнеКолонокКодовНеТрогается()
    {
        // E - это колонка подписей слева от кодов, к сумме по поставке отношения не имеет.
        Assert.Null(LoaderFormulaRepair.Expand("=SUM(E8:E8)", First, Last));
    }

    [Fact]
    public void СсылкаНаДругойЛистНеТрогается()
    {
        // На «Распреде» свои размеры, и его диапазоны Excel правит сам.
        Assert.Null(LoaderFormulaRepair.Expand(
            "=SUMPRODUCT(Распред!$I$28:$I$28)", First, Last));
    }

    [Fact]
    public void ЗнакиДоллараСохраняются()
    {
        Assert.Equal("=SUM($G$13:$I$13)", LoaderFormulaRepair.Expand("=SUM($I$13:$I$13)", First, Last));
    }

    [Fact]
    public void ЧинитсяКаждыйПодходящийДиапазонФормулы()
    {
        Assert.Equal(
            "=SUM(G8:I8)-SUM(G9:I9)",
            LoaderFormulaRepair.Expand("=SUM(I8:I8)-SUM(I9:I9)", First, Last));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("155588")]
    public void НеФормула_НеТрогается(string? formula)
    {
        Assert.Null(LoaderFormulaRepair.Expand(formula, First, Last));
    }

    [Fact]
    public void КолонкаОднаИПослеРасширения_ЧинитьНечего()
    {
        Assert.Null(LoaderFormulaRepair.Expand("=SUM(G13:G13)", First, First));
    }
}
