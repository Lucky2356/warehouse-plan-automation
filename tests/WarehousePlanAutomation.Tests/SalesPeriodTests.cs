using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class SalesPeriodTests
{
    [Theory]
    [InlineData("01/08-31/01", 1, 8, 31, 1)]
    [InlineData("25/08-31/01", 25, 8, 31, 1)]
    [InlineData(" 01/10 - 31/01 ", 1, 10, 31, 1)]
    [InlineData("01.10-31.01", 1, 10, 31, 1)]
    public void РазбираетПериодПродаж(string text, int startDay, int startMonth, int endDay, int endMonth)
    {
        Assert.True(SalesPeriodParser.TryParse(text, out var period));
        Assert.Equal(startDay, period.StartDay);
        Assert.Equal(startMonth, period.StartMonth);
        Assert.Equal(endDay, period.EndDay);
        Assert.Equal(endMonth, period.EndMonth);
    }

    [Theory]
    [InlineData("")]
    [InlineData("0")]
    [InlineData("01/08")]
    [InlineData("#Н/Д")]
    [InlineData("40/08-31/01")]
    [InlineData("01/13-31/01")]
    public void НеразбираемоеЗначениеНеПревращаетсяВПериод(string text)
    {
        Assert.False(SalesPeriodParser.TryParse(text, out _));
    }

    [Fact]
    public void СравнениеСДатамиИдётПоДнюИМесяцу()
    {
        // Года в тексте нет, поэтому и сравнение по нему не идёт: придумывать год
        // значило бы решать за книгу.
        Assert.True(SalesPeriodParser.TryParse("01/08-31/01", out var period));

        Assert.True(SalesPeriodParser.Matches(period, new DateTime(2026, 8, 1), new DateTime(2027, 1, 31)));
        Assert.True(SalesPeriodParser.Matches(period, new DateTime(2024, 8, 1), new DateTime(2025, 1, 31)));
        Assert.False(SalesPeriodParser.Matches(period, new DateTime(2026, 10, 1), new DateTime(2027, 1, 31)));
    }

    [Fact]
    public void ОбратноеПреобразованиеДаётТотЖеТекст()
    {
        var period = SalesPeriodParser.FromDates(new DateTime(2026, 8, 1), new DateTime(2027, 1, 31));

        Assert.Equal("01/08-31/01", period.ToString());
    }

    [Theory]
    [InlineData(CellError.NotAvailable)]
    [InlineData(CellError.DivideByZero)]
    [InlineData(CellError.Value)]
    public void КодыОшибокExcelУзнаютсяКакОшибки(int code)
    {
        Assert.True(CellError.IsError(code));
    }

    [Fact]
    public void ОбычныеЗначенияОшибкамиНеСчитаются()
    {
        Assert.False(CellError.IsError(0d));
        Assert.False(CellError.IsError(599d));
        Assert.False(CellError.IsError("01/08-31/01"));
        Assert.False(CellError.IsError(null));
    }

    [Theory]
    [InlineData("150,155,156,158,159,161")]
    [InlineData("1 234")]
    [InlineData("-12,5")]
    [InlineData("1E+5")]
    public void СтрокиКоторыеExcelПриметЗаЧисло_Узнаются(string text)
    {
        // Перечень магазинов из «Denny goods» Excel принимает за число с разделителями
        // разрядов и превращает в 1,5E+113. Такие строки пишутся в текстовом формате.
        Assert.True(CellError.LooksNumericToExcel(text));
    }

    [Theory]
    [InlineData("МЕЛКИЕ АКСЕССУАРЫ")]
    [InlineData("22х35см")]
    [InlineData("431-2802СЕРЫЙ-ЧЕРНЫЙTU")]
    [InlineData("TU")]
    [InlineData("")]
    [InlineData(null)]
    public void ОбычныйТекстЗаЧислоНеПринимается(string? text)
    {
        Assert.False(CellError.LooksNumericToExcel(text));
    }

    [Fact]
    public void НедоступноеЗначениеОтличаетсяОтПрочихОшибок()
    {
        Assert.True(CellError.IsNotAvailable(CellError.NotAvailable));
        Assert.True(CellError.IsNotAvailable("#Н/Д"));
        Assert.False(CellError.IsNotAvailable(CellError.DivideByZero));
    }
}
