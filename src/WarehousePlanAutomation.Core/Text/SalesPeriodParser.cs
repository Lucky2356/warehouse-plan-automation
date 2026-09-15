using System.Globalization;

namespace WarehousePlanAutomation.Core.Text;

/// <summary>Период продаж в том виде, в каком он записан в стенках: «01/08-31/01».</summary>
public readonly record struct SalesPeriod(int StartDay, int StartMonth, int EndDay, int EndMonth)
{
    public override string ToString() =>
        Two(StartDay) + "/" + Two(StartMonth) + "-" + Two(EndDay) + "/" + Two(EndMonth);

    private static string Two(int value) => value.ToString("00", CultureInfo.InvariantCulture);
}

/// <summary>
/// Разбор периода продаж. Года в тексте нет - только день и месяц, поэтому и сравнение
/// с датами листа «link» идёт по дню и месяцу. Придумывать год было бы догадкой.
/// </summary>
public static class SalesPeriodParser
{
    public static bool TryParse(string? text, out SalesPeriod period)
    {
        period = default;

        var normalized = TextUtils.Normalize(text).Replace(" ", string.Empty);
        var parts = normalized.Split('-', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2)
        {
            return false;
        }

        if (!TryParseDayMonth(parts[0], out var startDay, out var startMonth) ||
            !TryParseDayMonth(parts[1], out var endDay, out var endMonth))
        {
            return false;
        }

        period = new SalesPeriod(startDay, startMonth, endDay, endMonth);
        return true;
    }

    /// <summary>Совпадает ли период с парой дат листа «link». Сравниваются день и месяц.</summary>
    public static bool Matches(SalesPeriod period, DateTime start, DateTime end) =>
        period.StartDay == start.Day && period.StartMonth == start.Month &&
        period.EndDay == end.Day && period.EndMonth == end.Month;

    public static SalesPeriod FromDates(DateTime start, DateTime end) =>
        new(start.Day, start.Month, end.Day, end.Month);

    private static bool TryParseDayMonth(string text, out int day, out int month)
    {
        day = 0;
        month = 0;

        var parts = text.Split(new[] { '/', '.' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2)
        {
            return false;
        }

        return int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out day) &&
               int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out month) &&
               day is >= 1 and <= 31 &&
               month is >= 1 and <= 12;
    }
}
