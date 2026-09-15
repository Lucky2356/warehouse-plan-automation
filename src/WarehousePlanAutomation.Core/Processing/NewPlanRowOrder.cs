using System.Text.RegularExpressions;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Порядок новых строк «Плана» внутри блока - по правилам, которые задала аналитик:
/// 1. срочные («сроч» в «Поставках») выше несрочных;
/// 2. по дате в сети из текста: «в рознице с 01.10» выше «в рознице с 08.10»; строки без даты
///    («отгрузка по готовности») идут первыми, дата им не придумывается;
/// 3. при одной дате - сначала СЕТ, потом МОНО, потом остальное; СЕТ1 выше СЕТ2.
/// При полном равенстве остаётся порядок выгрузки. Старые строки блока не переставляются -
/// сортируются только сегодняшние.
/// </summary>
public static class NewPlanRowOrder
{
    private static readonly Regex SetPattern = new(@"(?<!\p{L})сет(?!\p{L})\s*(?<number>\d*)", RegexOptions.CultureInvariant);

    private static readonly Regex MonoPattern = new(@"(?<!\p{L})моно(?!\p{L})", RegexOptions.CultureInvariant);

    private const string UrgencyMarker = "сроч";

    public static IReadOnlyList<NewPlanRowSpec> Sort(IReadOnlyList<NewPlanRowSpec> rows) =>
        rows
            .OrderBy(row => TextUtils.ContainsKey(row.Supplies, UrgencyMarker) ? 0 : 1)
            .ThenBy(row => row.NetworkDate is null ? 0 : 1)
            .ThenBy(row => row.NetworkDate?.Date ?? DateTime.MinValue)
            .ThenBy(row => SetMonoRank(row.Supplies))
            .ThenBy(row => SetNumber(row.Supplies))
            .ToList();

    /// <summary>
    /// СЕТ - 0, МОНО - 1, остальное - 2. «СЕТ» и «МОНО» бывают набраны латиницей («CET2», «MOHO»).
    /// Слово ищется целиком: «в сети» и «кассета» признаком не считаются.
    /// </summary>
    public static int SetMonoRank(string? supplies)
    {
        var text = ToCyrillic(TextUtils.NormalizeKey(supplies));
        if (SetPattern.IsMatch(text))
        {
            return 0;
        }

        return MonoPattern.IsMatch(text) ? 1 : 2;
    }

    /// <summary>Номер СЕТ («СЕТ2» - 2) для порядка внутри СЕТ; без номера и не СЕТ - 0.</summary>
    private static int SetNumber(string? supplies)
    {
        var match = SetPattern.Match(ToCyrillic(TextUtils.NormalizeKey(supplies)));
        return match.Success && int.TryParse(match.Groups["number"].Value, out var number) ? number : 0;
    }

    /// <summary>Латинские буквы, похожие на русские, в словах «СЕТ» и «МОНО».</summary>
    private static string ToCyrillic(string text)
    {
        const string latin = "cetmoh";
        const string cyrillic = "сетмон";

        var chars = text.ToCharArray();
        for (var i = 0; i < chars.Length; i++)
        {
            var index = latin.IndexOf(chars[i]);
            if (index >= 0)
            {
                chars[i] = cyrillic[index];
            }
        }

        return new string(chars);
    }
}
