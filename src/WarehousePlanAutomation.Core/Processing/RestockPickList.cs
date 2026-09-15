using System.Globalization;
using System.Text;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Код товара на месте хранения: сколько на нём и из какой он поставки.</summary>
public sealed record StockCode(string Code, double Quantity, string SupplyNumber);

/// <summary>Почему у строки листа «из‹место›» другой код.</summary>
public enum PickMark
{
    /// <summary>Код, который находит ВПР по АЦР, и на нём хватает.</summary>
    None,

    /// <summary>На первом коде не хватает, взят другой код, где хватает. Строка голубая.</summary>
    ReplacedCode,

    /// <summary>Ни на одном коде не хватает - количество разнесено по нескольким. Строки зелёные.</summary>
    SplitCode,
}

/// <summary>Строка листа «из‹место›».</summary>
/// <param name="Shortage">Сколько не хватило на кодах этого места. Больше нуля - сообщить человеку.</param>
public sealed record PickLine(
    int RowIndex,
    StoragePlace Place,
    string Code,
    double Quantity,
    double CodeQuantity,
    string SupplyNumber,
    PickMark Mark,
    double Shortage = 0d);

/// <summary>
/// Код для строки листа «из‹место›». ВПР по АЦР находит первый код на листе места хранения;
/// если на нём столько нет, берётся код, на котором хватает (из «3, 1 и 30» - тот, где 30),
/// а если хватает только вместе - строка размножается по кодам.
/// </summary>
public static class PickListBuilder
{
    public static IReadOnlyList<PickLine> Build(
        int rowIndex,
        StoragePlace place,
        double quantity,
        IReadOnlyList<StockCode> codes)
    {
        if (codes.Count == 0)
        {
            return new[] { new PickLine(rowIndex, place, string.Empty, quantity, 0d, string.Empty, PickMark.None, quantity) };
        }

        var first = codes[0];
        if (first.Quantity >= quantity)
        {
            return new[] { Line(first, quantity, PickMark.None) };
        }

        var best = codes.Aggregate((a, b) => b.Quantity > a.Quantity ? b : a);
        if (best.Quantity >= quantity)
        {
            return new[] { Line(best, quantity, PickMark.ReplacedCode) };
        }

        var positive = codes
            .Select((code, order) => (code, order))
            .Where(item => item.code.Quantity > 0d)
            .OrderByDescending(item => item.code.Quantity)
            .ThenBy(item => item.order)
            .Select(item => item.code)
            .ToList();

        if (positive.Count == 0)
        {
            return new[] { Line(first, quantity, PickMark.None, quantity - Math.Max(first.Quantity, 0d)) };
        }

        var lines = new List<PickLine>();
        var remaining = quantity;
        for (var i = 0; i < positive.Count && remaining > 0d; i++)
        {
            var code = positive[i];
            var isLast = i == positive.Count - 1;
            var take = isLast ? remaining : Math.Min(code.Quantity, remaining);
            lines.Add(Line(code, take, PickMark.SplitCode, Math.Max(take - code.Quantity, 0d)));
            remaining -= take;
        }

        if (lines.Count == 1)
        {
            lines[0] = lines[0] with { Mark = ReferenceEquals(positive[0], first) ? PickMark.None : PickMark.ReplacedCode };
        }

        return lines;

        PickLine Line(StockCode code, double take, PickMark mark, double shortage = 0d) =>
            new(rowIndex, place, code.Code, take, code.Quantity, code.SupplyNumber, mark, shortage);
    }
}

/// <summary>Один загрузочник: строки одного места хранения с одним «коментом».</summary>
public sealed record RestockLoader(
    string SheetName,
    StoragePlace Place,
    string Comment,
    IReadOnlyList<PickLine> Lines);

/// <summary>
/// Загрузочники подтоварки. На каждое место хранения и каждый «комент» - свой лист:
/// «Очки отдельно» на «иза» и на «измп» - два разных загрузочника. Лист называется
/// «З‹место›-‹начало комента›»: «ЗМП-любой», «ЗА-мелкий», «ЗМПП-352».
/// </summary>
public static class RestockLoaderBuilder
{
    /// <summary>Самое длинное название листа, которое принимает Excel.</summary>
    public const int SheetNameLimit = 31;

    public static IReadOnlyList<RestockLoader> Build(
        IReadOnlyList<PickLine> lines,
        Func<int, string> commentOfRow)
    {
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<RestockLoader>();

        foreach (var place in StoragePlaces.Priority)
        {
            var groups = lines
                .Where(line => line.Place == place)
                .GroupBy(line => TextUtils.NormalizeKey(commentOfRow(line.RowIndex)))
                .ToList();

            foreach (var group in groups)
            {
                var comment = TextUtils.Normalize(commentOfRow(group.First().RowIndex));
                result.Add(new RestockLoader(SheetName(place, comment, taken), place, comment, group.ToList()));
            }
        }

        return result;
    }

    /// <summary>«ЗМП-любой»: начало «комента» - первое слово или число, «352-148ЧЕРНЫЙ» даёт «352».</summary>
    public static string SheetName(StoragePlace place, string comment, ISet<string> taken)
    {
        var start = new StringBuilder();
        foreach (var ch in TextUtils.Normalize(comment))
        {
            if (!char.IsLetterOrDigit(ch))
            {
                if (start.Length > 0)
                {
                    break;
                }

                continue;
            }

            if (start.Length > 0 && char.IsDigit(ch) != char.IsDigit(start[^1]))
            {
                break;
            }

            start.Append(char.ToLowerInvariant(ch));
        }

        var name = "З" + StoragePlaces.Code(place) + (start.Length > 0 ? "-" + start : string.Empty);
        if (name.Length > SheetNameLimit)
        {
            name = name[..SheetNameLimit];
        }

        var unique = name;
        for (var copy = 2; !taken.Add(unique); copy++)
        {
            var suffix = " " + copy.ToString(CultureInfo.InvariantCulture);
            unique = (name.Length + suffix.Length > SheetNameLimit ? name[..(SheetNameLimit - suffix.Length)] : name) + suffix;
        }

        return unique;
    }

    /// <summary>
    /// «Комментарий» загрузочника: маркетплейс, секторы, место хранения (у поставок - их номера
    /// без повторов), приоритет. «Lamoda Подтоварка Сумки, Обувь из С318-156, С2437-027 Приоритет к 17.09».
    /// </summary>
    public static string CommentText(
        string marketplace,
        IEnumerable<string> sectors,
        StoragePlace place,
        IEnumerable<string> supplyNumbers,
        IEnumerable<string> priorities)
    {
        var parts = new List<string>();
        if (marketplace.Length > 0)
        {
            parts.Add(marketplace);
        }

        parts.Add("Подтоварка");

        var sectorText = string.Join(", ", Distinct(sectors).Select(SentenceCase));
        if (sectorText.Length > 0)
        {
            parts.Add(sectorText);
        }

        parts.Add("из " + PlaceText(place, supplyNumbers));

        var priorityText = string.Join(", ", Distinct(priorities));
        if (priorityText.Length > 0)
        {
            parts.Add("Приоритет " + priorityText);
        }

        return string.Join(" ", parts);
    }

    /// <summary>
    /// «Приоритет» как его пишут в комментарии: текст - как есть («к 17.09»), дата - «к 07.09.2026».
    /// В одних книгах колонка текстовая, в других там дата, и в ячейке лежит число.
    /// </summary>
    public static string PriorityText(object? value)
    {
        if (value is double serial && serial is > 36526d and < 73051d)
        {
            return "к " + DateTime.FromOADate(Math.Floor(serial)).ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
        }

        return TextUtils.Normalize(TextUtils.CellToString(value));
    }

    public static string PlaceText(StoragePlace place, IEnumerable<string> supplyNumbers) => place switch
    {
        StoragePlace.Marketplace => "адресов МП",
        StoragePlace.Storage => "А1,А2,А3",
        StoragePlace.Returns => "Возвратов",
        _ => string.Join(", ", Distinct(supplyNumbers)),
    };

    /// <summary>
    /// Как маркетплейс называют в комментариях заказов: «Lamoda Подтоварка…», «Ozon Мск…».
    /// Для незнакомого кода клиента берётся хвост его группы подразделения с листа «Р»:
    /// «МеркетПлейсы / Вайлдберриз» → «Вайлдберриз».
    /// </summary>
    public static string MarketplaceName(string? clientCode, string? divisionGroup)
    {
        switch (TextUtils.NormalizeKey(clientCode))
        {
            case "208":
                return "Lamoda";
            case "206":
                return "Ozon";
            case "205":
                return "WB";
            case "211":
                return "Золотое Яблоко";
        }

        var group = TextUtils.Normalize(divisionGroup);
        var slash = group.LastIndexOf('/');
        return slash >= 0 ? group[(slash + 1)..].Trim() : group;
    }

    /// <summary>«ШАРФЫ И ПЛАТКИ» → «Шарфы и платки».</summary>
    public static string SentenceCase(string text)
    {
        var normalized = TextUtils.Normalize(text).ToLowerInvariant();
        return normalized.Length == 0
            ? normalized
            : char.ToUpperInvariant(normalized[0]) + normalized[1..];
    }

    private static IEnumerable<string> Distinct(IEnumerable<string> values)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var value in values)
        {
            var normalized = TextUtils.Normalize(value);
            if (normalized.Length > 0 && seen.Add(TextUtils.NormalizeKey(normalized)))
            {
                yield return normalized;
            }
        }
    }
}

/// <summary>Какие строки «Р» остаются: этот год, текущий месяц и два предыдущих.</summary>
public static class ReservePeriod
{
    /// <summary>
    /// «Если сегодня сентябрь 2026 - из месяцев оставим только июль, август и сентябрь».
    /// Прошлые годы удаляются всегда, поэтому в январе остаётся только январь.
    /// </summary>
    public static bool Keep(DateTime date, DateTime today) =>
        date.Year >= today.Year && (date.Year > today.Year || date.Month >= today.Month - 2);
}
