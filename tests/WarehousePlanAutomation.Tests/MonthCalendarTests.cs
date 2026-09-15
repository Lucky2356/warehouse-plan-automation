using WarehousePlanAutomation.Core.Text;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class MonthCalendarTests
{
    [Fact]
    public void ОктябрьДвадцатьШестого_ЭтоНедели40По44()
    {
        // Проверено по производственному календарю: неделя 40 начинается 28 сентября,
        // и в октябрь от неё попадают четыре дня - этого хватает.
        Assert.Equal(new[] { 40, 41, 42, 43, 44 }, MonthCalendar.WeeksOf(new YearMonth(2026, 10)));
    }

    [Fact]
    public void НоябрьДвадцатьШестого_ЭтоНедели45По48()
    {
        // От недели 44 в ноябре только 1 ноября, от недели 49 - только 30-е:
        // такие недели в расчёт не идут.
        Assert.Equal(new[] { 45, 46, 47, 48 }, MonthCalendar.WeeksOf(new YearMonth(2026, 11)));
    }

    [Fact]
    public void ДекабрьДвадцатьШестого_ЗаканчиваетсяПервойНеделей()
    {
        // Неделя 28 декабря захватывает 1 января, поэтому в календаре она первая -
        // и колонка «1» на листе «КС» именно эта. В декабре её дней четыре, она считается.
        Assert.Equal(new[] { 49, 50, 51, 52, 1 }, MonthCalendar.WeeksOf(new YearMonth(2026, 12)));
    }

    [Fact]
    public void ЯнварьДвадцатьСедьмого_НачинаетсяСПервойНедели()
    {
        // В январе от недели 28 декабря три дня - ровно столько, сколько нужно.
        Assert.Equal(new[] { 1, 2, 3, 4, 5 }, MonthCalendar.WeeksOf(new YearMonth(2027, 1)));
    }

    [Fact]
    public void НеделяСДвумяДнямиВМесяцНеПопадает()
    {
        // Май 2026 начинается в пятницу: от недели 27 апреля в мае три дня, она считается.
        // А в апреле от недели 30 марта пять дней, значит эта неделя попадает в оба месяца.
        Assert.Contains(18, MonthCalendar.WeeksOf(new YearMonth(2026, 5)));
        Assert.Contains(18, MonthCalendar.WeeksOf(new YearMonth(2026, 4)));

        // Август 2026 начинается в субботу: от недели 27 июля в августе два дня - не в счёт.
        Assert.DoesNotContain(31, MonthCalendar.WeeksOf(new YearMonth(2026, 8)));
    }

    [Fact]
    public void НомерНеделиСчитаетсяОтНеделиПервогоЯнваря()
    {
        Assert.Equal(1, MonthCalendar.WeekNumber(new DateTime(2026, 1, 1)));
        Assert.Equal(1, MonthCalendar.WeekNumber(new DateTime(2025, 12, 29)));
        Assert.Equal(40, MonthCalendar.WeekNumber(new DateTime(2026, 10, 1)));
        Assert.Equal(52, MonthCalendar.WeekNumber(new DateTime(2026, 12, 27)));
        Assert.Equal(1, MonthCalendar.WeekNumber(new DateTime(2026, 12, 28)));
    }

    [Fact]
    public void ТриМесяцаИдутПодряд()
    {
        var months = MonthCalendar.ThreeFrom(new YearMonth(2026, 11));

        Assert.Equal(new[] { "Ноябрь", "Декабрь", "Январь" }, months.Select(m => m.Name));
        Assert.Equal(new[] { 2026, 2026, 2027 }, months.Select(m => m.Year));
    }

    [Fact]
    public void СезонностьСчитаетсяСоСледующегоМесяца()
    {
        var months = MonthCalendar.SeasonFrom(new DateTime(2026, 9, 9));

        Assert.Equal(
            new[] { "Октябрь 2026", "Ноябрь 2026", "Декабрь 2026" },
            months.Select(m => m.ToString()));
    }

    [Fact]
    public void СезонностьВДекабреПереходитНаСледующийГод()
    {
        var months = MonthCalendar.SeasonFrom(new DateTime(2026, 12, 20));

        Assert.Equal(
            new[] { "Январь 2027", "Февраль 2027", "Март 2027" },
            months.Select(m => m.ToString()));
    }

    [Theory]
    [InlineData("Октябрь 2026", 2026, 10)]
    [InlineData("  декабрь   2027 ", 2027, 12)]
    public void РазбираетНазваниеМесяцаСГодом(string text, int year, int month)
    {
        Assert.True(MonthCalendar.TryParse(text, out var parsed));
        Assert.Equal(new YearMonth(year, month), parsed);
    }

    [Theory]
    [InlineData("Октябрь")]
    [InlineData("Октябрь 2026 года")]
    [InlineData("Октябрину 2026")]
    [InlineData("")]
    public void НеразбираемыйОтветНеПревращаетсяВМесяц(string text)
    {
        Assert.False(MonthCalendar.TryParse(text, out _));
    }
}
