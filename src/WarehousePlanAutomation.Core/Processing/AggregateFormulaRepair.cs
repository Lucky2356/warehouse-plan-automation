using System.Globalization;
using System.Text.RegularExpressions;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Восстановление диапазона итоговой формулы блока после вставки и удаления строк.
///
/// Смысл исходной формулы сохраняется: если вчера итог считался не с первой строки блока
/// (в блоке «все группы» сумма начинается после строк «Приемка на хранилище» и
/// «автозаказы для хабов»), то и после обработки сохраняется тот же отступ от начала блока.
/// </summary>
public static class AggregateFormulaRepair
{
    private static readonly Regex SimpleSumPattern = new(
        @"^=\s*SUM\(\s*(?<col1>\$?[A-Z]{1,3})\$?(?<row1>\d+)\s*:\s*(?<col2>\$?[A-Z]{1,3})\$?(?<row2>\d+)\s*\)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>
    /// Возвращает исправленную формулу или null, если формула не является простой суммой
    /// одного диапазона: в этом случае её трогать нельзя.
    /// </summary>
    public static string? BuildRepairedFormula(
        string? originalFormula,
        int originalFirstDataRow,
        int newFirstDataRow,
        int newLastDataRow)
    {
        if (string.IsNullOrWhiteSpace(originalFormula) || newLastDataRow < newFirstDataRow)
        {
            return null;
        }

        var match = SimpleSumPattern.Match(originalFormula.Trim());
        if (!match.Success)
        {
            return null;
        }

        var column1 = match.Groups["col1"].Value;
        var column2 = match.Groups["col2"].Value;
        if (!string.Equals(column1, column2, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var startRow = int.Parse(match.Groups["row1"].Value, CultureInfo.InvariantCulture);
        var offset = startRow - originalFirstDataRow;
        if (offset < 0)
        {
            offset = 0;
        }

        var newStart = newFirstDataRow + offset;
        if (newStart > newLastDataRow)
        {
            newStart = newLastDataRow;
        }

        return "=SUM(" + column1 + newStart.ToString(CultureInfo.InvariantCulture) +
               ":" + column2 + newLastDataRow.ToString(CultureInfo.InvariantCulture) + ")";
    }
}
