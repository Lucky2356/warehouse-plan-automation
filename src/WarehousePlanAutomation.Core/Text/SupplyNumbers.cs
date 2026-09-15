using System.Text.RegularExpressions;

namespace WarehousePlanAutomation.Core.Text;

/// <summary>
/// Номера поставок: «С352-136», «М318-157», «Л318-151». Буква говорит, что это за поставка,
/// цифры - сама поставка.
///
/// В «Удалили из плана склада» и в «Плане склада» номера пишутся без буквы, через запятую
/// и часто сокращённо: «2445-011,014,015» - это 2445-011, 2445-014 и 2445-015,
/// «327-037,38» - 327-037 и 327-038. Сокращение продолжает последний полный номер:
/// короткое число заменяет его последние цифры.
/// </summary>
public static class SupplyNumbers
{
    /// <summary>Буква поставок маркетплейса.</summary>
    public const char Marketplace = 'м';

    /// <summary>Буква обычных поставок сети.</summary>
    public const char Network = 'с';

    /// <summary>Буква поставок, которые из разбора убираются.</summary>
    public const char Excluded = 'л';

    private static readonly Regex Full = new(
        @"(?<!\d)(\d{2,5})-(\d{2,4})(?!\d)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex Continuation = new(
        @"\G\s*,\s*(\d{1,4})(?![\d-])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>Цифры номера без буквы: «С352-136» → «352-136». Пусто, если номера нет.</summary>
    public static string Digits(string? number)
    {
        var match = Full.Match(number ?? string.Empty);
        return match.Success ? match.Groups[1].Value + "-" + match.Groups[2].Value : string.Empty;
    }

    /// <summary>
    /// Буква перед номером, строчная кириллица. Латинские двойники («C», «M») читаются
    /// как кириллические: на глаз их не отличить, а в выгрузках встречаются. '\0' - буквы нет.
    /// </summary>
    public static char Letter(string? number)
    {
        foreach (var raw in TextUtils.Normalize(number))
        {
            if (char.IsDigit(raw))
            {
                break;
            }

            if (!char.IsLetter(raw))
            {
                continue;
            }

            return char.ToLowerInvariant(raw) switch
            {
                'c' => Network,
                'm' => Marketplace,
                var letter => letter == 'ё' ? 'е' : letter,
            };
        }

        return '\0';
    }

    /// <summary>Все номера поставок в тексте, сокращённые - развёрнутыми.</summary>
    public static IReadOnlySet<string> Extract(string? text)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(text))
        {
            return result;
        }

        foreach (Match match in Full.Matches(text))
        {
            var prefix = match.Groups[1].Value;
            var suffix = match.Groups[2].Value;
            result.Add(prefix + "-" + suffix);

            var position = match.Index + match.Length;
            while (true)
            {
                var next = Continuation.Match(text, position);
                if (!next.Success)
                {
                    break;
                }

                var token = next.Groups[1].Value;
                if (token.Length > suffix.Length)
                {
                    break;
                }

                result.Add(prefix + "-" + suffix[..(suffix.Length - token.Length)] + token);
                position = next.Index + next.Length;
            }
        }

        return result;
    }

    /// <summary>Номера из нескольких текстов сразу - например, из всех ячеек листа.</summary>
    public static IReadOnlySet<string> ExtractAll(IEnumerable<string?> texts)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (var text in texts)
        {
            result.UnionWith(Extract(text));
        }

        return result;
    }

    /// <summary>Ключ полного номера для сравнения: буква и цифры, «С352-136» → «с352-136».</summary>
    public static string Key(string? number)
    {
        var digits = Digits(number);
        if (digits.Length == 0)
        {
            return string.Empty;
        }

        var letter = Letter(number);
        return letter == '\0' ? digits : letter + digits;
    }
}
