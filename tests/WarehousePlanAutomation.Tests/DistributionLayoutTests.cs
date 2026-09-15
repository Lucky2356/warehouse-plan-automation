using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class DistributionLayoutTests
{
    /// <summary>
    /// Синтетический «Распред», повторяющий разметку реального: блок сезонности в F1:J14,
    /// подписи в колонке M, блоки АЦР по четыре колонки начиная с N, таблица РТТ с 27-й строки.
    /// </summary>
    private static SheetGrid BuildGrid(int blocks = 3, int stores = 3)
    {
        const int columns = 40;
        var rowCount = 27 + stores;
        var values = new object?[rowCount, columns];
        var formulas = new string?[rowCount, columns];

        void Set(int row, int column, object? value) => values[row - 1, column - 1] = value;
        void Formula(int row, int column, string text) => formulas[row - 1, column - 1] = text;

        Set(1, 6, "Сезонность по НК текущего сезона");
        Set(1, 7, "Октябрь");
        Set(1, 10, "Мин раздача");
        Set(2, 6, "МАСКА ДЛЯ СНА/");

        Set(6, 13, "мин запас на Хаб при наличии первички");
        Set(18, 13, "Код");
        Set(22, 13, "Ед");
        Set(26, 13, "Ост");
        Set(26, 8, "МЕЛКИЕ АКСЕССУАРЫ");

        Set(27, 1, "Код");
        Set(27, 11, "рейтинг");
        Set(27, 13, "Итого");
        for (var block = 0; block < blocks; block++)
        {
            var column = 14 + (block * 4);
            Set(27, column, "распред");
            Set(27, column + 1, "запас на ХАБ");
            Set(27, column + 2, "Остатки");
            Set(27, column + 3, "в загрузку");
        }

        for (var store = 0; store < stores; store++)
        {
            var row = 28 + store;
            Set(row, 1, 18 + store);
            Formula(row, 4, "=VLOOKUP(TEXT($A" + row + ",\"000\")&$H$26,'типы по секторам'!A:J,10,0)");
            Formula(row, 14, "=IF($B" + row + "=\"ХАБ\",0,VLOOKUP(N$2,$F$2:$J$14,2,0))");
        }

        return new SheetGrid(1, 1, values, formulas);
    }

    [Fact]
    public void НаходитБлокиИТаблицуРТТ()
    {
        var layout = DistributionLayout.Read(BuildGrid());

        Assert.Equal(27, layout.HeaderRow);
        Assert.Equal(13, layout.LabelColumn);
        Assert.Equal(14, layout.FirstBlockColumn);
        Assert.Equal(3, layout.BlockCount);
        Assert.Equal(25, layout.LastBlockColumn);
        Assert.Equal(28, layout.StoreFirstRow);
        Assert.Equal(30, layout.StoreLastRow);
    }

    [Fact]
    public void НаходитПодписиСлеваОтБлоков()
    {
        var layout = DistributionLayout.Read(BuildGrid());

        Assert.Equal(18, layout.CodeRow);
        Assert.Equal(22, layout.UnitsRow);
        Assert.Equal(26, layout.RemainderRow);
        Assert.Equal(6, layout.HubMinimumRow);
    }

    [Fact]
    public void БерётБлокСезонностиИзФормулыРаспределения()
    {
        // Блок нигде не подписан: его адрес известен только формуле «распред».
        var layout = DistributionLayout.Read(BuildGrid());

        Assert.Equal(new RangeRef(2, 6, 14, 10), layout.Seasonality);
        Assert.Equal(1, layout.MonthRow);
        Assert.Equal(6, layout.SubgroupColumn);
        Assert.Equal(3, layout.MonthCount);
        Assert.Equal(7, layout.MonthColumn(0));
        Assert.Equal(13, layout.SubgroupCapacity);
    }

    [Fact]
    public void БерётЯчейкуСектораИзФормулыТаблицыРТТ()
    {
        var layout = DistributionLayout.Read(BuildGrid());

        Assert.Equal(new CellRef(26, 8), layout.SectorCell);
    }

    [Fact]
    public void БлокиАдресуютсяПоНомеру()
    {
        var layout = DistributionLayout.Read(BuildGrid());

        Assert.Equal(14, layout.BlockColumn(0));
        Assert.Equal(18, layout.BlockColumn(1));
        Assert.Equal(22, layout.BlockColumn(2));
    }

    [Fact]
    public void ОдинБлокТожеРазметка()
    {
        var layout = DistributionLayout.Read(BuildGrid(blocks: 1));

        Assert.Equal(1, layout.BlockCount);
        Assert.Equal(17, layout.LastBlockColumn);
    }

    [Fact]
    public void БезСтрокиЗаголовков_ПонятнаяОшибка()
    {
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]> { new object?[] { "что-то" } });

        var exception = Assert.Throws<WorkbookValidationException>(() => DistributionLayout.Read(grid));
        Assert.Contains("в загрузку", exception.Message);
    }
}
