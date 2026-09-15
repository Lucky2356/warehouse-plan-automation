using System.Globalization;
using System.Text.RegularExpressions;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Переводит ссылки на лист «остатки» на его фактические размеры.
///
/// Формулы «Распред» написаны под прошлую выгрузку: «остатки!$C$14:$NZ$14»,
/// «остатки!$C$35:$NZ$393». Число АЦР и число строк меняются с каждой поставкой, и в
/// инструкции это отдельный шаг - «смотрим адрес самой последней колонки и меняем его
/// в двух формулах». Excel сам такие ссылки не растягивает: лист другой.
///
/// Правило простое и не зависит от того, что именно считает формула: у диапазона,
/// растянутого по колонкам, меняется последняя колонка; у растянутого по строкам -
/// последняя строка. Диапазон в одну колонку или одну строку остаётся таким же.
/// </summary>
public static class StockReferenceRepair
{
    private static readonly Regex RangePattern = new(
        @"(?<sheet>'[^']+'|[^\s!'"",();*+/\-=<>&]+)!" +
        @"\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>\d{1,7})\s*:\s*\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>\d{1,7})",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Возвращает исправленную формулу или null, если менять нечего.
    /// </summary>
    public static string? Retarget(string? formula, string sheetName, int lastRow, int lastColumn)
    {
        if (string.IsNullOrWhiteSpace(formula) ||
            !formula.StartsWith("=", StringComparison.Ordinal) ||
            lastRow < 1 ||
            lastColumn < 1)
        {
            return null;
        }

        var changed = false;
        var result = RangePattern.Replace(formula, match =>
        {
            var sheet = match.Groups["sheet"].Value.Trim('\'');
            if (!Text.TextUtils.EqualsKey(sheet, Text.TextUtils.NormalizeKey(sheetName)))
            {
                return match.Value;
            }

            var firstColumn = ExcelColumn.FromLetters(match.Groups["c1"].Value);
            var secondColumn = ExcelColumn.FromLetters(match.Groups["c2"].Value);
            var firstRow = int.Parse(match.Groups["r1"].Value, CultureInfo.InvariantCulture);
            var secondRow = int.Parse(match.Groups["r2"].Value, CultureInfo.InvariantCulture);

            var targetColumn = firstColumn == secondColumn ? secondColumn : lastColumn;
            var targetRow = firstRow == secondRow ? secondRow : lastRow;

            if (targetColumn == secondColumn && targetRow == secondRow)
            {
                return match.Value;
            }

            changed = true;
            return match.Groups["sheet"].Value + "!" +
                   "$" + match.Groups["c1"].Value + "$" + firstRow.ToString(CultureInfo.InvariantCulture) +
                   ":" +
                   "$" + ExcelColumn.ToLetters(targetColumn) +
                   "$" + targetRow.ToString(CultureInfo.InvariantCulture);
        });

        return changed ? result : null;
    }
}
