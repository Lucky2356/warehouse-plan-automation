using System.Text.RegularExpressions;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Починка сумм по колонкам кодов на листе «Загрузочник».
///
/// Обычно такие суммы Excel растягивает сам: вставка колонки внутрь диапазона
/// «СУММ(G13:I13)» расширяет его. Но диапазон из одной ячейки расширить нельзя -
/// «СУММ(G13:G13)» при вставке просто уезжает вправо и продолжает считать одну колонку.
/// Так бывает после поставки с единственным кодом: её загрузочник становится заготовкой
/// для следующей, и суммы в ней молча считают не то.
///
/// Правило узкое нарочно: чинится только диапазон, у которого совпадают и строка,
/// и колонка, то есть одна ячейка, записанная диапазоном, и стоит она среди колонок
/// кодов. Сумма по колонке «СУММ(G13:G131)» под него не подходит - у неё разные
/// строки - и остаётся как есть.
/// </summary>
public static class LoaderFormulaRepair
{
    private static readonly Regex RangePattern = new(
        @"(?<d1>\$?)(?<c1>[A-Za-z]{1,3})(?<e1>\$?)(?<r1>\d{1,7})\s*:\s*(?<d2>\$?)(?<c2>[A-Za-z]{1,3})(?<e2>\$?)(?<r2>\d{1,7})",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Возвращает исправленную формулу или null, если менять нечего.
    /// </summary>
    public static string? Expand(string? formula, int firstColumn, int lastColumn)
    {
        if (string.IsNullOrEmpty(formula) || formula[0] != '=' || firstColumn >= lastColumn)
        {
            return null;
        }

        var changed = false;

        var repaired = RangePattern.Replace(formula, match =>
        {
            if (IsQualified(formula, match.Index))
            {
                return match.Value;
            }

            var column1 = ExcelColumn.FromLetters(match.Groups["c1"].Value);
            var column2 = ExcelColumn.FromLetters(match.Groups["c2"].Value);
            var row1 = match.Groups["r1"].Value;
            var row2 = match.Groups["r2"].Value;

            var single = column1 == column2 && row1 == row2;
            var inBand = column1 >= firstColumn && column1 <= lastColumn;

            if (!single || !inBand)
            {
                return match.Value;
            }

            changed = true;
            return match.Groups["d1"].Value + ExcelColumn.ToLetters(firstColumn) +
                   match.Groups["e1"].Value + row1 + ":" +
                   match.Groups["d2"].Value + ExcelColumn.ToLetters(lastColumn) +
                   match.Groups["e2"].Value + row2;
        });

        return changed ? repaired : null;
    }

    /// <summary>Перед ссылкой стоит «!» - значит, она относится к другому листу.</summary>
    private static bool IsQualified(string formula, int index)
    {
        for (var i = index - 1; i >= 0; i--)
        {
            var symbol = formula[i];
            if (symbol == '!')
            {
                return true;
            }

            if (!char.IsLetterOrDigit(symbol) && symbol != '$' && symbol != '\'' &&
                symbol != ' ' && symbol != '_' && symbol != '.')
            {
                return false;
            }
        }

        return false;
    }
}
