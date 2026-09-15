using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Выбранная для подгруппы строка регламента и то, как её выбрали.</summary>
public sealed record MarkupChoice(string Sector, string Subgroup, MarkupRule? Rule, string? Warning);

/// <summary>
/// Выбор наценок.
///
/// Если сектор встречается в регламенте один раз - выбирать нечего.
/// Если несколько («СУМКИ рюкзак / мешок / кошелек / прочее»), строки различаются колонкой
/// «Сектор+», и по инструкции строка выбирается по подгруппе товара, а если по подгруппе
/// не нашлась - берётся «Прочее». Спрашивает программа только тогда, когда и так ответа нет:
/// у «ОБУВЬ зима / лето / деми» ни подгруппа, ни «Прочее» выбора не дают.
///
/// Выбор делается один раз на подгруппу: в одной поставке подгрупп может быть несколько,
/// и наценки у них тогда тоже разные.
/// </summary>
public sealed class MarkupResolver
{
    private readonly IReadOnlyList<MarkupRule> _rules;
    private readonly IDecisionPrompt? _prompt;
    private readonly Dictionary<string, MarkupChoice> _resolved = new(StringComparer.Ordinal);

    public MarkupResolver(IReadOnlyList<MarkupRule> rules, IDecisionPrompt? prompt = null)
    {
        _rules = rules;
        _prompt = prompt;
    }

    public MarkupChoice Resolve(string? sector, string? subgroup = null)
    {
        var key = TextUtils.NormalizeKey(sector);
        var subgroupName = TextUtils.Normalize(subgroup);

        if (key.Length == 0)
        {
            return new MarkupChoice(
                string.Empty, subgroupName, null, "У товара не заполнен сектор, наценки не проставлены.");
        }

        var cacheKey = key + "|" + TextUtils.NormalizeKey(subgroup);
        if (_resolved.TryGetValue(cacheKey, out var cached))
        {
            return cached;
        }

        var choice = Choose(TextUtils.Normalize(sector), subgroupName, key);
        _resolved[cacheKey] = choice;
        return choice;
    }

    private MarkupChoice Choose(string sector, string subgroup, string key)
    {
        var matches = _rules
            .Where(rule => TextUtils.EqualsKey(rule.Sector, key))
            .ToList();

        if (matches.Count == 0)
        {
            return new MarkupChoice(
                sector,
                subgroup,
                null,
                "Сектор «" + sector + "» не найден в «" + PriceSchema.MarkupSheet + "», наценки не проставлены.");
        }

        if (matches.Count == 1)
        {
            return new MarkupChoice(sector, subgroup, matches[0], null);
        }

        var candidates = MarkupSubgroupMatch.Candidates(sector, subgroup, matches);
        if (candidates.Count == 1)
        {
            return new MarkupChoice(sector, subgroup, candidates[0], null);
        }

        var about = subgroup.Length > 0
            ? "подгруппы «" + subgroup + "» (сектор «" + sector + "»)"
            : "сектора «" + sector + "»";

        var options = candidates.Select(Describe).ToList();
        var answer = _prompt?.Choose(new DecisionRequest(
            "Какую наценку взять для " + about + "?",
            "В «" + PriceSchema.MarkupSheet + "» для этого сектора несколько строк, они различаются " +
            "колонкой «Сектор+». По подгруппе строка не определилась, а строки «Прочее» " +
            (candidates.Count == matches.Count ? "у сектора нет" : "несколько") +
            ". Если подгрупп в поставке несколько, вопрос будет по каждой.",
            options));

        if (answer is null)
        {
            return new MarkupChoice(
                sector,
                subgroup,
                null,
                "Для " + about + " в регламенте несколько строк (" +
                string.Join("; ", candidates.Select(m => m.SectorPlus)) +
                "), выбор не сделан - наценки не проставлены.");
        }

        var index = options.FindIndex(option => string.Equals(option, answer, StringComparison.Ordinal));
        return index < 0
            ? new MarkupChoice(sector, subgroup, null, "Ответ «" + answer + "» не совпал ни с одной строкой регламента.")
            : new MarkupChoice(sector, subgroup, candidates[index], null);
    }

    private static string Describe(MarkupRule rule)
    {
        var name = rule.SectorPlus.Length > 0 ? rule.SectorPlus : rule.Sector;
        return name +
               "  —  плановая " + Format(rule.Planned) +
               ", минимальная " + Format(rule.Minimum);
    }

    private static string Format(double? value) =>
        value is null ? "не задана" : value.Value.ToString("0.##", System.Globalization.CultureInfo.CurrentCulture);
}

/// <summary>
/// Строка регламента по подгруппе товара.
///
/// «Сектор+» пишется сокращениями и слитно: «МЕЛ АКСЕСС под упак», «ГОЛ УБОРЫ зима»,
/// «ШАРФЫ И ПЛАТКИлето». От него отбрасываются слова самого сектора («мел» - начало
/// «мелкие», «аксесс» - начало «аксессуары»), и остаются слова, которыми строка
/// отличается: «под упак», «зима». Подгруппа совпадает со строкой, если в её названии
/// есть каждое такое слово - целиком или началом («кошелек» - «КОШЕЛЬКИ»).
///
/// Строка без отличающих слов («ПЕРЧАТКИ» у сектора «ПЕРЧАТКИ») и строка со словом
/// «прочее» («ПРОЧЕЕ», «СУМКИ прочее») - общие: они берутся, когда подгруппа
/// ни с одной отдельной строкой не совпала.
/// </summary>
public static class MarkupSubgroupMatch
{
    private const string OtherPrefix = "проч";

    /// <summary>Общее начало, начиная с которого разные окончания одного слова считаются одним словом.</summary>
    private const int StemLength = 5;

    private static readonly HashSet<string> Conjunctions = new(StringComparer.Ordinal) { "и", "для", "под" };

    /// <summary>
    /// Строки, из которых остаётся выбрать. Одна - выбор сделан. Несколько - решает человек:
    /// в этом случае возвращаются совпавшие по подгруппе строки, иначе общие, иначе все.
    /// </summary>
    public static IReadOnlyList<MarkupRule> Candidates(string sector, string subgroup, IReadOnlyList<MarkupRule> rules)
    {
        var sectorWords = Words(sector);
        var subgroupWords = Words(subgroup);

        var specific = new List<MarkupRule>();
        var general = new List<MarkupRule>();

        foreach (var rule in rules)
        {
            var words = Distinctive(rule, sectorWords);
            if (words.Count == 0 || words.Any(word => word.StartsWith(OtherPrefix, StringComparison.Ordinal)))
            {
                general.Add(rule);
                continue;
            }

            if (subgroupWords.Count > 0 &&
                words.All(word => subgroupWords.Any(candidate => SameWord(word, candidate))))
            {
                specific.Add(rule);
            }
        }

        if (specific.Count > 0)
        {
            return specific;
        }

        return general.Count > 0 ? general : rules;
    }

    /// <summary>
    /// Слова «Сектор+», которых нет в названии сектора. Служебные слова («и», «для», «под»)
    /// ничего не различают и тоже отбрасываются: «под» нашлось бы в любой «подвеске».
    /// </summary>
    private static List<string> Distinctive(MarkupRule rule, IReadOnlyList<string> sectorWords)
    {
        var text = rule.SectorPlus.Length > 0 ? rule.SectorPlus : rule.Sector;
        return Words(text)
            .Where(word => !Conjunctions.Contains(word))
            .Where(word => !sectorWords.Any(sectorWord =>
                sectorWord.StartsWith(word, StringComparison.Ordinal) ||
                word.StartsWith(sectorWord, StringComparison.Ordinal)))
            .ToList();
    }

    private static bool SameWord(string word, string candidate)
    {
        if (candidate.StartsWith(word, StringComparison.Ordinal) || word.StartsWith(candidate, StringComparison.Ordinal))
        {
            return Math.Min(word.Length, candidate.Length) >= 3;
        }

        var common = 0;
        while (common < word.Length && common < candidate.Length && word[common] == candidate[common])
        {
            common++;
        }

        return common >= StemLength;
    }

    /// <summary>
    /// Слова в нижнем регистре. Граница слова - всё, что не буква, и переход от прописной
    /// к строчной: «ПЛАТКИлето» - это два слова.
    /// </summary>
    public static IReadOnlyList<string> Words(string? text)
    {
        var words = new List<string>();
        var value = TextUtils.Normalize(text);
        var current = new System.Text.StringBuilder();
        var previousUpper = false;

        void Flush()
        {
            if (current.Length > 0)
            {
                words.Add(current.ToString().ToLowerInvariant().Replace('ё', 'е'));
                current.Clear();
            }
        }

        foreach (var ch in value)
        {
            if (!char.IsLetter(ch))
            {
                Flush();
                previousUpper = false;
                continue;
            }

            var upper = char.IsUpper(ch);
            if (current.Length > 1 && previousUpper && !upper)
            {
                Flush();
            }

            current.Append(ch);
            previousUpper = upper;
        }

        Flush();
        return words;
    }
}
