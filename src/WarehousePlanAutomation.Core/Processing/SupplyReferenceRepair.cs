namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Ссылки листа «Цены» на лист «для цен». Каждая поставка смотрит в свою колонку:
/// первая - в «B», вторая - в «C». В книге это ссылки вида ='для цен'!$B$5/'для цен'!$B$7,
/// и меняется в них только буква колонки.
///
/// Перевод нужен в обе стороны: в прошлом файле поставок могло быть две, а в этом одна -
/// тогда строка, смотревшая в «C», должна вернуться на «B».
/// </summary>
public static class SupplyReferenceRepair
{
    /// <summary>
    /// Переводит все абсолютные ссылки формулы на колонку <paramref name="columnLetters"/>.
    /// Возвращает null, если менять нечего.
    /// </summary>
    public static string? Retarget(string? formula, string columnLetters)
    {
        if (string.IsNullOrEmpty(formula) || columnLetters.Length == 0)
        {
            return null;
        }

        var result = new System.Text.StringBuilder(formula.Length);
        var changed = false;
        var index = 0;

        while (index < formula.Length)
        {
            if (formula[index] != '$')
            {
                result.Append(formula[index]);
                index++;
                continue;
            }

            var letters = index + 1;
            while (letters < formula.Length && char.IsAsciiLetter(formula[letters]))
            {
                letters++;
            }

            // Колонка узнаётся по второму «$»: $B$5 - это колонка, а $5 или $B - нет.
            if (letters == index + 1 || letters >= formula.Length || formula[letters] != '$')
            {
                result.Append(formula[index]);
                index++;
                continue;
            }

            var current = formula[(index + 1)..letters];
            changed |= !string.Equals(current, columnLetters, StringComparison.OrdinalIgnoreCase);
            result.Append('$').Append(columnLetters).Append('$');
            index = letters + 1;
        }

        return changed ? result.ToString() : null;
    }
}
