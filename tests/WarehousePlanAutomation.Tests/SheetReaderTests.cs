using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class SheetReaderTests
{
    [Fact]
    public void ВыгрузкаЗаказов_КолонкаРазницаЕдНаходитсяПоНачалуЗаголовка()
    {
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[]
            {
                "Дата документа", "Дата план", "Дата закрытия", "Номер", "Подразделение", "Комментарий",
                "Кол-во док", "Поз,Арт", "Кол-во фкт", "Поз,Арт,Мест", "Разница ед", "Статус",
            },
            new object?[]
            {
                46210.7d, 46210d, "---", "З000-254935", "Ozon",
                "Ozon Мск Подтоварка Номер загрузки 54750654", 64d, 61010d, 0d, "000", 64d, "ЗАПУЩЕН",
            },
        });

        var sheet = OrdersSheetReader.Read(grid);

        Assert.Equal(11, sheet.Headers[SheetSchema.Orders.DifferenceUnits]);
        Assert.Single(sheet.Rows);
        Assert.Equal("Ozon", sheet.Rows[0].Division);
        Assert.Equal(64d, sheet.Rows[0].DifferenceUnits);
        Assert.Equal(46210.7d, sheet.Rows[0].DocumentDate);
    }

    [Fact]
    public void Журнал_КолонкаПроцентаНаходитсяТолькоПоТочномуСовпадению()
    {
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[]
            {
                "Направление", "Статус", "Комментарий", "Выбран", "Город", "Подразделение", "Номер",
                "Дата план", "Шк внтары", "Дата закрытия", "Кол-во ед", "Факт ед", "Разница",
                "%", "% оклейка", "% отгрузка", "% сборка",
            },
            new object?[]
            {
                "Магазины M", "ЗАКРЫТ", "Подтоварка Номер загрузки 55575395", 0d, "-17", "-32", "-32",
                46262d, null, null, 15952d, 13189d, 2763d, 83d, 100d, 100d, 100d,
            },
        });

        var sheet = JournalSheetReader.Read(grid);

        Assert.Equal(14, sheet.Headers[SheetSchema.Journal.Percent]);
        Assert.Equal(2, sheet.Headers[SheetSchema.Journal.Status]);
        Assert.Single(sheet.Rows);
        Assert.Equal(83d, sheet.Rows[0].Percent);
        Assert.Equal(0, sheet.Rows[0].Order);
    }

    [Fact]
    public void Журнал_СохраняетФизическийПорядокСтрок()
    {
        var rows = new List<object?[]>
        {
            new object?[] { "Направление", "Статус", "Комментарий", "Номер", "Кол-во ед", "Факт ед", "%" },
        };

        for (var i = 0; i < 5; i++)
        {
            rows.Add(new object?[]
            {
                "Магазины M", "ЗАПУЩЕН", "Номер загрузки 5557539" + i, "З000-26035" + i, 100d, (double)i, (double)i,
            });
        }

        var sheet = JournalSheetReader.Read(SheetGrid.FromRows(1, 1, rows));

        Assert.Equal(new[] { 0, 1, 2, 3, 4 }, sheet.Rows.Select(r => r.Order));
        Assert.Equal(new[] { 2, 3, 4, 5, 6 }, sheet.Rows.Select(r => r.ExcelRow));
    }

    [Fact]
    public void ExcelColumn_ПреобразуетНомераВБуквы()
    {
        Assert.Equal("A", ExcelColumn.ToLetters(1));
        Assert.Equal("Z", ExcelColumn.ToLetters(26));
        Assert.Equal("AA", ExcelColumn.ToLetters(27));
        Assert.Equal("Q", ExcelColumn.ToLetters(17));
        Assert.Equal(17, ExcelColumn.FromLetters("Q"));
        Assert.Equal(27, ExcelColumn.FromLetters("aa"));
    }
}
