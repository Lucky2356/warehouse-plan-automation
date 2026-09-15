using System.Globalization;

namespace WarehousePlanAutomation.Core.Text;

/// <summary>Месяц с годом: год нужен для номеров недель.</summary>
public readonly record struct YearMonth(int Year, int Month)
{
    public YearMonth Next() => Month == 12 ? new YearMonth(Year + 1, 1) : new YearMonth(Year, Month + 1);

    /// <summary>Название месяца в именительном падеже и с большой буквы: «Октябрь».</summary>
    public string Name => MonthCalendar.Names[Month - 1];

    public override string ToString() => Name + " " + Year.ToString(CultureInfo.InvariantCulture);
}

/// <summary>
/// Месяцы и номера недель - так, как их считает производственный календарь.
///
/// Первая неделя года - та, на которую приходится 1 января; недели идут с понедельника.
/// По этому счёту в 2026 году 52 недели, ровно столько колонок и на листе «КС»,
/// а неделя 28 декабря - это уже первая неделя следующего года.
///
/// Неделя относится к месяцу, если в этом месяце лежат хотя бы три её дня: неделю,
/// от которой в месяц попал один-два дня, в расчёт сезонности не берут. Пограничная
/// неделя может попасть сразу в два месяца - так её и считают вручную.
/// </summary>
public static class MonthCalendar
{
    /// <summary>Сколько дней недели должно попасть в месяц, чтобы неделя пошла в расчёт.</summary>
    public const int MinimumDaysInMonth = 3;

    public static readonly IReadOnlyList<string> Names = new[]
    {
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
    };

    /// <summary>Понедельник недели, в которую попадает эта дата.</summary>
    public static DateTime MondayOf(DateTime date) =>
        date.Date.AddDays(-(((int)date.DayOfWeek + 6) % 7));

    /// <summary>
    /// Номер недели по производственному календарю: неделя с 1 января - первая.
    /// Неделя, захватившая 1 января следующего года, получает номер 1 - так она
    /// и подписана в календаре, и такая колонка есть на листе «КС».
    /// </summary>
    public static int WeekNumber(DateTime date)
    {
        var day = date.Date;
        if (day >= FirstMonday(day.Year + 1))
        {
            return 1;
        }

        return ((day - FirstMonday(day.Year)).Days / 7) + 1;
    }

    /// <summary>
    /// Номера недель, у которых в этом месяце не меньше <see cref="MinimumDaysInMonth"/> дней.
    /// </summary>
    public static IReadOnlyList<int> WeeksOf(YearMonth month)
    {
        var days = DateTime.DaysInMonth(month.Year, month.Month);
        var counts = new Dictionary<DateTime, int>();

        for (var day = 1; day <= days; day++)
        {
            var monday = MondayOf(new DateTime(month.Year, month.Month, day));
            counts[monday] = counts.TryGetValue(monday, out var current) ? current + 1 : 1;
        }

        return counts
            .Where(pair => pair.Value >= MinimumDaysInMonth)
            .OrderBy(pair => pair.Key)
            .Select(pair => WeekNumber(pair.Key))
            .ToList();
    }

    /// <summary>Месяц и два следующих за ним.</summary>
    public static IReadOnlyList<YearMonth> ThreeFrom(YearMonth first)
    {
        var second = first.Next();
        return new[] { first, second, second.Next() };
    }

    /// <summary>
    /// Три месяца сезонности: считают на месяцы вперёд, а текущий уже наполовину прошёл,
    /// поэтому отсчёт идёт со следующего. Про сезонность программа больше не спрашивает -
    /// какие месяцы взяты, она пишет в замечаниях.
    /// </summary>
    public static IReadOnlyList<YearMonth> SeasonFrom(DateTime today) =>
        ThreeFrom(new YearMonth(today.Year, today.Month).Next());

    public static bool TryParse(string? text, out YearMonth month)
    {
        month = default;

        var parts = TextUtils.Normalize(text).Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2 || !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var year))
        {
            return false;
        }

        for (var i = 0; i < Names.Count; i++)
        {
            if (TextUtils.EqualsKey(parts[0], Names[i].ToLowerInvariant()))
            {
                month = new YearMonth(year, i + 1);
                return true;
            }
        }

        return false;
    }

    /// <summary>Понедельник первой недели года: недели с 1 января.</summary>
    private static DateTime FirstMonday(int year) => MondayOf(new DateTime(year, 1, 1));
}
