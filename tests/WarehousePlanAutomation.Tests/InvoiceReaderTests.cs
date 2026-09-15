using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class InvoiceReaderTests
{
    /// <summary>
    /// Повторяет разметку реального инвойса: шапка с номером, строка заголовков
    /// по-русски, сразу под ней такая же по-английски, три товара и строка «итого».
    /// </summary>
    private static SheetGrid BuildGrid() => SheetGrid.FromRows(1, 1, new List<object?[]>
    {
        new object?[] { null, "INVOICE", null, null, null },
        new object?[] { null, "Инвойс № и дата", null, null, null },
        new object?[] { null, "431-329", null, null, null },
        new object?[] { null, "номер модели", "штрих-код", "количество", "цена", "сумма" },
        new object?[] { null, "Item No.", "Barcode", "Quantity", "price, (CNY)", "Amount, (CNY)" },
        new object?[] { null, "SLEEP MASK", 1431280202d, 1010d, 6.55d, 6615.5d },
        new object?[] { null, "SLEEP MASK", 1431333901d, 840d, 6.55d, 5502d },
        new object?[] { null, "SLEEP MASK", 1431334001d, 815d, 6.55d, 5338.25d },
        new object?[] { null, null, null, 2665d, null, 17455.75d },
    });

    [Fact]
    public void ЧитаетТолькоСтрокиТовара()
    {
        var invoice = InvoiceSheetReader.Read(BuildGrid(), "Invoice");

        Assert.Equal(3, invoice.Lines.Count);
        Assert.Equal("1431280202", invoice.Lines[0].Barcode);
        Assert.Equal(1010d, invoice.Lines[0].Quantity);
        Assert.Equal(6.55d, invoice.Lines[0].Price);
        Assert.Equal(6615.5d, invoice.Lines[0].Amount);
    }

    [Fact]
    public void АнглийскаяСтрокаЗаголовковНеСчитаетсяТоваром()
    {
        // Под русскими заголовками стоит вторая строка заголовков: штрихкод в ней есть,
        // но «количество» - слово, а не число.
        var invoice = InvoiceSheetReader.Read(BuildGrid(), "Invoice");

        Assert.DoesNotContain(invoice.Lines, line => line.Barcode.Contains("Barcode"));
    }

    [Fact]
    public void СтрокаИтогоНеСчитаетсяТоваром()
    {
        var invoice = InvoiceSheetReader.Read(BuildGrid(), "Invoice");

        Assert.Equal(2665d, invoice.TotalQuantity);
        Assert.Equal(17455.75d, invoice.TotalAmount);
    }

    [Fact]
    public void НомерПоставкиБерётсяИзШапки()
    {
        Assert.Equal("431-329", InvoiceSheetReader.Read(BuildGrid(), "Invoice").SupplyNumber);
    }

    [Fact]
    public void БезПодписиВШапкеНомерПоставкиПустой()
    {
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[] { "штрих-код", "количество", "цена", "сумма" },
            new object?[] { 1431280202d, 10d, 1d, 10d },
        });

        Assert.Equal(string.Empty, InvoiceSheetReader.Read(grid, "Invoice").SupplyNumber);
    }

    [Fact]
    public void БезНомераВШапке_НомерПоЛистуИФайлу()
    {
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[] { "штрих-код", "количество", "цена", "сумма" },
            new object?[] { 1431280202d, 10d, 1d, 10d },
        });

        Assert.Equal("431-341", InvoiceSheetReader.Read(
            grid, "Invoice-341", "431-338, 431-341 Бижутерия LTL722_до").SupplyNumber);
    }

    [Theory]
    [InlineData("Invoice-338", "431-338, 431-341 Бижутерия", "431-338")]
    [InlineData("Invoice (341)", "431-338, 431-341 Бижутерия", "431-341")]
    [InlineData("Invoice", "431-338, 431-341 Бижутерия", "")]
    [InlineData("Invoice-400", "431-338, 431-341 Бижутерия", "")]
    [InlineData("Invoice-338", "431-338, 432-338 Бижутерия", "")]
    public void НомерПоставкиПоНазваниям(string sheet, string file, string expected)
    {
        Assert.Equal(expected, InvoiceSheetReader.SupplyFromNames(sheet, file));
    }

    [Fact]
    public void НомерВШапкеСильнееНазваний()
    {
        Assert.Equal("431-329", InvoiceSheetReader.Read(BuildGrid(), "Invoice-341", "431-338, 431-341").SupplyNumber);
    }

    [Fact]
    public void ЗаголовокСВалютойСчитаетсяТемЖе()
    {
        // В инвойсах поставщиков к названию приписана валюта: «цена, CNY», «сумма, CNY».
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[] { "модель", "штрих-код", "количество", "цена, CN¥", "сумма, CNY" },
            new object?[] { "SLEEP MASK", 1431280202d, 1010d, 6.55d, 6615.5d },
        });

        var line = Assert.Single(InvoiceSheetReader.Read(grid, "Invoice").Lines);

        Assert.Equal(6.55d, line.Price);
        Assert.Equal(6615.5d, line.Amount);
    }

    [Fact]
    public void ЕдиницаИзмеренияВСкобкахТожеОтбрасывается()
    {
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[] { "штрих-код", "количество (шт.)", "цена (CNY)", "сумма (CNY)" },
            new object?[] { 1431280202d, 1010d, 6.55d, 6615.5d },
        });

        Assert.Equal(1010d, Assert.Single(InvoiceSheetReader.Read(grid, "Invoice").Lines).Quantity);
    }

    [Fact]
    public void ВалютаБезЗапятойТожеОтбрасывается()
    {
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[] { "штрих-код", "количество", "цена CN¥", "сумма CNY" },
            new object?[] { 1431280202d, 1010d, 6.55d, 6615.5d },
        });

        var line = Assert.Single(InvoiceSheetReader.Read(grid, "Invoice").Lines);

        Assert.Equal(6.55d, line.Price);
        Assert.Equal(6615.5d, line.Amount);
    }

    [Fact]
    public void ПродолжениеНазванияНеСчитаетсяЕдиницейИзмерения()
    {
        // Отбрасывается только хвост после запятой или скобки: «цена продажи» - это
        // другая колонка, и подставлять её вместо «цены» нельзя.
        var grid = SheetGrid.FromRows(1, 1, new List<object?[]>
        {
            new object?[] { "штрих-код", "количество", "цена продажи", "сумма" },
            new object?[] { 1431280202d, 1010d, 6.55d, 6615.5d },
        });

        var error = Assert.Throws<WorkbookValidationException>(() => InvoiceSheetReader.Read(grid, "Invoice"));

        Assert.Contains(error.Problems, problem => problem.Contains("«цена»"));
    }

    [Fact]
    public void ДлинныйШтрихкодНеПревращаетсяВЭкспоненту()
    {
        // Через double штрихкод прочитался бы как «1,9E+09» и перестал бы
        // совпадать со «Сводным прайсом».
        Assert.Equal("1431280202", InvoiceSheetReader.NormalizeBarcode("1431280202"));
        Assert.Equal("1431280202", InvoiceSheetReader.NormalizeBarcode(" 1431 280 202 "));
    }
}
