using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Подгруппа со своей группой: в «КС» одна и та же подгруппа встречается в разных группах.</summary>
public sealed record SubgroupKey(string Group, string Subgroup);

/// <summary>Строка блока сезонности: подгруппа и доля продаж по каждому из месяцев.</summary>
public sealed record SeasonalityShare(SubgroupKey Key, IReadOnlyList<double?> Shares);

/// <summary>Что записать в блок сезонности листа «Распред».</summary>
public sealed class SeasonalityPlan
{
    public SeasonalityPlan(
        IReadOnlyList<YearMonth> months,
        IReadOnlyList<SeasonalityShare> rows,
        IReadOnlyList<ProcessingWarning> warnings)
    {
        Months = months;
        Rows = rows;
        Warnings = warnings;
    }

    public IReadOnlyList<YearMonth> Months { get; }

    public IReadOnlyList<SeasonalityShare> Rows { get; }

    public IReadOnlyList<ProcessingWarning> Warnings { get; }
}

/// <summary>
/// Доля продаж подгруппы за месяц - сумма её недельных долей с листа «КС».
/// Недели берутся по производственному календарю: первая неделя года - та, на которую
/// приходится 1 января, и неделя идёт в расчёт месяца, если в этом месяце лежат хотя бы
/// три её дня. Пограничная неделя может попасть сразу в два месяца - так её и считают
/// вручную. Правило целиком - в <see cref="MonthCalendar"/>.
///
/// Строка «КС» ищется по паре «Группа + Подгруппа»: одна и та же подгруппа встречается
/// в разных группах. Если под пару подходит несколько строк с разными числами, программа
/// не выбирает - пишет замечание и оставляет ячейки пустыми.
/// </summary>
public static class SeasonalityCalculator
{
    public static SeasonalityPlan Build(
        IReadOnlyList<SubgroupKey> subgroups,
        IReadOnlyList<YearMonth> months,
        IReadOnlyList<SeasonalityRow> seasonality)
    {
        var warnings = new List<ProcessingWarning>();
        var rows = new List<SeasonalityShare>(subgroups.Count);
        var reportedMissingWeeks = new HashSet<string>(StringComparer.Ordinal);

        foreach (var key in subgroups)
        {
            var matches = seasonality
                .Where(row => TextUtils.EqualsKey(row.Subgroup, TextUtils.NormalizeKey(key.Subgroup)) &&
                              TextUtils.EqualsKey(row.Group, TextUtils.NormalizeKey(key.Group)))
                .ToList();

            if (matches.Count == 0)
            {
                warnings.Add(new ProcessingWarning(
                    "Подгруппа «" + key.Subgroup + "» (группа «" + key.Group + "») не найдена в «" +
                    PriceSchema.SeasonalitySheet + "»: доли по месяцам не заполнены.",
                    PriceSchema.SeasonalitySheet));
                rows.Add(new SeasonalityShare(key, months.Select(_ => (double?)null).ToList()));
                continue;
            }

            if (matches.Count > 1 && !AllSame(matches))
            {
                warnings.Add(new ProcessingWarning(
                    "Для подгруппы «" + key.Subgroup + "» (группа «" + key.Group + "») в «" +
                    PriceSchema.SeasonalitySheet + "» " + matches.Count +
                    " строки с разными долями - программа не выбирает, доли по месяцам не заполнены.",
                    PriceSchema.SeasonalitySheet));
                rows.Add(new SeasonalityShare(key, months.Select(_ => (double?)null).ToList()));
                continue;
            }

            var source = matches[0];
            var shares = new List<double?>(months.Count);

            foreach (var month in months)
            {
                var weeks = MonthCalendar.WeeksOf(month);
                var missing = weeks.Where(week => !source.Weeks.ContainsKey(week)).ToList();

                if (missing.Count > 0)
                {
                    var note = month.ToString() + ":" + string.Join(",", missing);
                    if (reportedMissingWeeks.Add(note))
                    {
                        warnings.Add(new ProcessingWarning(
                            "В «" + PriceSchema.SeasonalitySheet + "» нет колонки для недели " +
                            string.Join(", ", missing) + " - доля за " + month.Name.ToLowerInvariant() +
                            " посчитана без неё и занижена.",
                            PriceSchema.SeasonalitySheet));
                    }
                }

                shares.Add(weeks.Sum(week => source.Weeks.TryGetValue(week, out var value) ? value : 0d));
            }

            rows.Add(new SeasonalityShare(key, shares));
        }

        return new SeasonalityPlan(months, rows, warnings);
    }

    /// <summary>Повторы одной подгруппы безопасны, пока числа в них совпадают.</summary>
    private static bool AllSame(IReadOnlyList<SeasonalityRow> rows)
    {
        var first = rows[0].Weeks;
        return rows.Skip(1).All(row =>
            row.Weeks.Count == first.Count &&
            first.All(pair => row.Weeks.TryGetValue(pair.Key, out var value) &&
                              Math.Abs(value - pair.Value) < 1e-9));
    }
}
