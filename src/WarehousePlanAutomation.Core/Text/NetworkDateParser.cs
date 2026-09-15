using System.Globalization;
using System.Text.RegularExpressions;

namespace WarehousePlanAutomation.Core.Text;

/// <summary>Дата, найденная в тексте: где она стоит и что означает.</summary>
/// <param name="Start">Позиция первого символа даты в тексте, с нуля.</param>
/// <param name="Length">Сколько символов занимает дата.</param>
public readonly record struct TextDate(int Start, int Length, DateTime Date)
{
    /// <summary>Дата в виде числа Excel - так её и пишут в ячейку.</summary>
    public double OADate => Date.ToOADate();
}

/// <summary>
/// Дата в сети из текста поставки: «2437-029 Кеды на Юг СЕТ1_в рознице с 15.09»,
/// «Пуховики_получение в рознице 20.09», «Перчатки_в сети с 07.09».
///
/// Дату, совпадающую с датой документа, программа за дату в сети не принимает:
/// в «Срочная подтоварка 28.08_Хранение» 28.08 - день, когда заказ завели, и «Дата в сети»
/// у таких строк в плане считается формулой от даты документа.
/// </summary>
public static class NetworkDateParser
{
    /// <summary>
    /// «15.09», «1.10», «01.08.2026». Месяц - всегда две цифры: так дата не путается
    /// с числом вроде «2.5», а в реальных текстах плана месяц пишут именно так.
    /// </summary>
    private static readonly Regex DatePattern = new(
        @"(?<![\d.,])(?<day>\d{1,2})\.(?<month>\d{2})(?:\.(?<year>\d{4}|\d{2}))?(?![\d,]|\.\d)",
        RegexOptions.CultureInvariant);

    public static TextDate? Find(string? text, double? documentDate, DateTime today)
    {
        if (string.IsNullOrEmpty(text))
        {
            return null;
        }

        var reference = documentDate is { } oa ? DateTime.FromOADate(Math.Floor(oa)) : today.Date;

        foreach (Match match in DatePattern.Matches(text))
        {
            var day = int.Parse(match.Groups["day"].Value, CultureInfo.InvariantCulture);
            var month = int.Parse(match.Groups["month"].Value, CultureInfo.InvariantCulture);
            if (month is < 1 or > 12 || day < 1)
            {
                continue;
            }

            if (documentDate is not null && day == reference.Day && month == reference.Month)
            {
                continue;
            }

            var date = match.Groups["year"].Success
                ? Build(ParseYear(match.Groups["year"].Value), month, day)
                : Closest(reference, month, day);

            if (date is not null)
            {
                return new TextDate(match.Index, match.Length, date.Value);
            }
        }

        return null;
    }

    /// <summary>
    /// Года в тексте обычно нет. Берётся тот, при котором дата ближе всего к дате документа:
    /// «15.01» у заказа от 20 декабря - это январь следующего года, а не прошлого.
    /// </summary>
    private static DateTime? Closest(DateTime reference, int month, int day)
    {
        DateTime? best = null;
        for (var year = reference.Year - 1; year <= reference.Year + 1; year++)
        {
            var candidate = Build(year, month, day);
            if (candidate is null)
            {
                continue;
            }

            if (best is null ||
                Math.Abs((candidate.Value - reference).TotalDays) < Math.Abs((best.Value - reference).TotalDays))
            {
                best = candidate;
            }
        }

        return best;
    }

    private static int ParseYear(string text)
    {
        var year = int.Parse(text, CultureInfo.InvariantCulture);
        return text.Length == 2 ? 2000 + year : year;
    }

    private static DateTime? Build(int year, int month, int day) =>
        day <= DateTime.DaysInMonth(year, month) ? new DateTime(year, month, day) : null;
}
