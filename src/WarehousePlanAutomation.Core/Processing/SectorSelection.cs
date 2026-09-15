using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Один файл распреда - один сектор. В инвойсе сектор не написан: он появляется только
/// на листе «Цены», когда строки подтянуты из «Сводного прайса». Если секторов оказалось
/// несколько, нужный узнаётся по названию файла - именно так поставка и называется:
/// «C2518-084 Бижутерия LTL719.xlsx».
/// </summary>
public static class SectorSelection
{
    /// <summary>
    /// Слова короче этого в поиске не участвуют: «для» и «и» есть в половине названий
    /// и совпали бы с чем угодно.
    /// </summary>
    private const int MinWordLength = 4;

    /// <summary>Секторы строк по убыванию числа строк. Пустые сектора не считаются.</summary>
    public static IReadOnlyList<string> Sectors(IEnumerable<PriceRowValues> rows) =>
        rows
            .Select(row => TextUtils.Normalize(row.Reference?.Sector))
            .Where(sector => sector.Length > 0)
            .GroupBy(sector => sector, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key, StringComparer.CurrentCulture)
            .Select(group => group.Key)
            .ToList();

    /// <summary>
    /// Секторы, которые узнаются в названии файла. Сектор подходит, если хотя бы одно
    /// его слово стоит в названии отдельным словом: «КОЛГОТКИ,НОСКИ» находится по файлу
    /// «318-140 Носки FTL 36», а «УКРАШЕНИЯ ДЛЯ ВОЛОС» - по слову «украшения».
    /// </summary>
    public static IReadOnlyList<string> MatchFileName(IReadOnlyList<string> sectors, string? filePath)
    {
        var name = TextUtils.NormalizeKey(SafeFileName(filePath));
        if (name.Length == 0)
        {
            return Array.Empty<string>();
        }

        return sectors.Where(sector => Words(sector).Any(word => TextUtils.ContainsWord(name, word))).ToList();
    }

    private static string SafeFileName(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return string.Empty;
        }

        try
        {
            return Path.GetFileNameWithoutExtension(filePath);
        }
        catch (ArgumentException)
        {
            return filePath;
        }
    }

    private static IEnumerable<string> Words(string sector)
    {
        var key = TextUtils.NormalizeKey(sector);
        var word = new System.Text.StringBuilder();

        foreach (var symbol in key)
        {
            if (char.IsLetterOrDigit(symbol))
            {
                word.Append(symbol);
                continue;
            }

            if (word.Length >= MinWordLength)
            {
                yield return word.ToString();
            }

            word.Clear();
        }

        if (word.Length >= MinWordLength)
        {
            yield return word.ToString();
        }
    }
}
