using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class LoaderLayoutTests
{
    /// <summary>
    /// Синтетический «Загрузочник», повторяющий разметку реального: блок формул над
    /// таблицей, ячейка проверки в D9, строка заголовков в 12-й строке, «Код» в колонке F,
    /// дальше колонки кодов и «Итого».
    /// </summary>
    private static SheetGrid BuildGrid(int codes = 3, int stores = 3)
    {
        var totalColumn = 6 + codes + 1;
        var rowCount = 12 + stores;
        var values = new object?[rowCount, totalColumn];
        var formulas = new string?[rowCount, totalColumn];

        void Set(int row, int column, object? value) => values[row - 1, column - 1] = value;
        void Formula(int row, int column, string text) => formulas[row - 1, column - 1] = text;

        Set(3, 6, "Подгруппа");
        Set(4, 6, "АЦР");
        Set(7, 6, "Заказ");
        Set(8, 6, "поставка");
        Set(9, 6, "распред");
        Set(10, 6, "остаток");

        Formula(8, 5, "=SUM(G8:I8)");
        Formula(9, 5, "=SUM(G9:I9)");
        Formula(9, 4, "=E9=Распред!L25-Распред!M147");
        Formula(10, 5, "=E9/E8");

        Set(12, 3, "Концепт");
        Set(12, 4, "Доля");
        Set(12, 5, "Хаб");
        Set(12, 6, "Код");
        for (var code = 0; code < codes; code++)
        {
            Set(12, 7 + code, 155588 + code);
        }

        Set(12, totalColumn, "Итого");

        for (var store = 0; store < stores; store++)
        {
            var row = 13 + store;
            Set(row, 6, "1" + store + "b");
            Formula(row, totalColumn, "=SUM(G" + row + ":I" + row + ")");
        }

        return new SheetGrid(1, 1, values, formulas);
    }

    [Fact]
    public void НаходитСтрокуЗаголовковКолонкиКодовИТаблицу()
    {
        var layout = LoaderLayout.Read(BuildGrid());

        Assert.Equal(12, layout.HeaderRow);
        Assert.Equal(6, layout.CodeColumn);
        Assert.Equal(7, layout.FirstCodeColumn);
        Assert.Equal(3, layout.CodeColumnCount);
        Assert.Equal(9, layout.LastCodeColumn);
        Assert.Equal(10, layout.TotalColumn);
        Assert.Equal(13, layout.StoreFirstRow);
        Assert.Equal(15, layout.StoreLastRow);
    }

    [Fact]
    public void ЯчейкаПроверкиНаходитсяПоВидуФормулы()
    {
        // Подписи у неё нет: от соседей она отличается только тем, что это сравнение.
        var layout = LoaderLayout.Read(BuildGrid());

        Assert.Equal(new CellRef(9, 4), layout.CheckCell);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(5)]
    [InlineData(12)]
    public void ЧислоКолонокКодовЧитаетсяПоРасстояниюДоИтого(int codes)
    {
        var layout = LoaderLayout.Read(BuildGrid(codes));

        Assert.Equal(codes, layout.CodeColumnCount);
        Assert.Equal(7 + codes, layout.TotalColumn);
    }

    [Fact]
    public void ПослеВставкиКолонокРазметкаСдвигаетИтого()
    {
        var layout = LoaderLayout.Read(BuildGrid()).WithCodeColumnCount(5);

        Assert.Equal(5, layout.CodeColumnCount);
        Assert.Equal(11, layout.LastCodeColumn);
        Assert.Equal(12, layout.TotalColumn);
        Assert.Equal(new CellRef(9, 4), layout.CheckCell);
    }

    [Fact]
    public void НетСтрокиЗаголовков_ПонятноеСообщение()
    {
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[] { "Концепт", "Доля", "Хаб" },
        });

        var error = Assert.Throws<WorkbookValidationException>(() => LoaderLayout.Read(grid));

        Assert.Contains("не найдена строка заголовков", error.Message);
    }

    [Theory]
    [InlineData("=E9=Распред!L25-Распред!M147", true)]
    [InlineData("=A1>B1", true)]
    [InlineData("=SUM(G8:I8)", false)]
    [InlineData("=VLOOKUP(G12,Цены!$D:$BI,58,0)", false)]
    [InlineData("=IF(A1=B1,1,0)", false)]
    [InlineData("=CONCAT(\"a=b\")", false)]
    [InlineData("155588", false)]
    [InlineData(null, false)]
    public void СравнениеОтличаетсяОтОбычнойФормулы(string? formula, bool expected)
    {
        // Знак сравнения внутри скобок принадлежит функции, а не ячейке проверки.
        Assert.Equal(expected, LoaderLayout.IsComparison(formula));
    }
}
