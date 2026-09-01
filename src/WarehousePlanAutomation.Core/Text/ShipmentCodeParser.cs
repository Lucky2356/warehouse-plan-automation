using System.Text.RegularExpressions;

namespace WarehousePlanAutomation.Core.Text;

/// <summary>
/// Разбор номеров поставок вида 1022-015, 1244-001, МЗ806-133.
/// Формат: необязательные буквы (кириллица или латиница, любой регистр) + цифры + "-" + цифры.
/// Минимальная длина каждой числовой части — три знака: это отсекает служебные фрагменты
/// сезонности вида "FW26-27", которые в реальных выгрузках номером поставки не являются.
/// </summary>
public static class ShipmentCodeParser
{
    private static readonly Regex CodePattern = new(
        @"(?<![\p{L}\p{N}])(?:\p{L}{1,3})?(?<head>\d{3,})-(?<tail>\d{3,})(?![\d-])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static bool ContainsCode(string? text) =>
        !string.IsNullOrWhiteSpace(text) && CodePattern.IsMatch(TextUtils.Normalize(text));

    /// <summary>
    /// Возвращает нормализованные номера поставок ("цифры-цифры", без буквенного префикса)
    /// в порядке появления, без повторов.
    /// </summary>
    public static IReadOnlyList<string> ExtractCodes(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return Array.Empty<string>();
        }

        var result = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match match in CodePattern.Matches(TextUtils.Normalize(text)))
        {
            var code = match.Groups["head"].Value + "-" + match.Groups["tail"].Value;
            if (seen.Add(code))
            {
                result.Add(code);
            }
        }

        return result;
    }
}
