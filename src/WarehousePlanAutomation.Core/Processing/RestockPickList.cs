using System.Globalization;
using System.Text;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Код товара на месте хранения: сколько на нём и из какой он поставки.</summary>
public sealed record StockCode(string Code, double Quantity, string SupplyNumber);

/// <summary>Почему у строки листа «из‹место›» другой код.</summary>
public enum PickMark
{
    /// <summary>Наименьший код АЦР, и на нём хватает.</summary>
    None,

    /// <summary>На наименьшем коде не хватает, взят следующий, где хватает. Строка голубая.</summary>
    ReplacedCode,

    /// <summary>Ни на одном коде не хватает - количество разнесено по нескольким. Строки зелёные.</summary>
    SplitCode,
}

/// <summary>Строка листа «из‹место›».</summary>
/// <param name="Shortage">Сколько не хватило на кодах этого места. Больше нуля - сообщить человеку.</param>
/// <param name="City">Город, которому идёт это количество. Пусто - город один.</param>
public sealed record PickLine(
    int RowIndex,
    StoragePlace Place,
    string Code,
    double Quantity,
    double CodeQuantity,
    string SupplyNumber,
    PickMark Mark,
    double Shortage = 0d,
    string City = "");

/// <summary>
/// Код для строки листа «из‹место›». Коды АЦР перебираются с наименьшего: на котором
/// хватает, тот и берём. Если не хватает ни на одном - количество набирается по кодам
/// в том же порядке.
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

        var ordered = Ascending(codes);
        var first = ordered[0];
        if (first.Quantity >= quantity)
        {
            return new[] { Line(first, quantity, PickMark.None) };
        }

        // Следующий по счёту код, на котором хватает: искать самый большой не нужно -
        // берём ближайший подходящий.
        var enough = ordered.FirstOrDefault(code => code.Quantity >= quantity);
        if (enough is not null)
        {
            return new[] { Line(enough, quantity, PickMark.ReplacedCode) };
        }

        var positive = ordered.Where(code => code.Quantity > 0d).ToList();

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

    /// <summary>
    /// Коды от наименьшего к большему. Коды товара - числа, но записаны текстом, поэтому
    /// сравниваются как числа; всё, что числом не читается, уходит в конец по алфавиту.
    /// </summary>
    private static IReadOnlyList<StockCode> Ascending(IReadOnlyList<StockCode> codes) => codes
        .Select((code, order) => (code, order, number: Number(code.Code)))
        .OrderBy(item => item.number is null)
        .ThenBy(item => item.number ?? 0d)
        .ThenBy(item => item.code.Code, StringComparer.Ordinal)
        .ThenBy(item => item.order)
        .Select(item => item.code)
        .ToList();

    private static double? Number(string code) =>
        double.TryParse(
            TextUtils.Normalize(code), NumberStyles.Any, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;

    /// <summary>
    /// Раздача собранного по городам: города идут по очереди, каждый берёт своё количество
    /// из того, что осталось. Предыдущий город для следующего - всё равно что резерв.
    /// Городов нет - строки остаются как есть.
    /// </summary>
    public static IReadOnlyList<PickLine> SplitByCity(
        IReadOnlyList<PickLine> lines, IReadOnlyList<(string City, double Quantity)> cities)
    {
        if (cities.Count == 0 || lines.Count == 0)
        {
            return lines;
        }

        var demand = cities.Select(city => Math.Max(city.Quantity, 0d)).ToArray();

        // Всё, что не разошлось по городам (например, количество урезал разбор «Запрета»),
        // достаётся последнему городу: терять его нельзя.
        var collected = lines.Sum(line => line.Quantity);
        var planned = demand.Sum();
        if (planned < collected)
        {
            demand[^1] += collected - planned;
        }

        var result = new List<PickLine>();
        var index = 0;
        foreach (var line in lines)
        {
            var remaining = line.Quantity;
            var first = true;
            while (remaining > 0d && index < demand.Length)
            {
                if (demand[index] <= 0d)
                {
                    index++;
                    continue;
                }

                var take = Math.Min(demand[index], remaining);
                result.Add(line with
                {
                    Quantity = take,
                    City = cities[index].City,
                    Shortage = first ? line.Shortage : 0d,
                });

                demand[index] -= take;
                remaining -= take;
                first = false;
            }

            if (remaining > 0d)
            {
                result.Add(line with { Quantity = remaining, City = cities[^1].City, Shortage = first ? line.Shortage : 0d });
            }
        }

        return result;
    }
}

/// <summary>Один загрузочник: строки одного места хранения с одним «коментом».</summary>
/// <param name="Place">Место хранения. null - загрузочник собран по всем местам сразу.</param>
public sealed record RestockLoader(
    string SheetName,
    StoragePlace? Place,
    string Comment,
    IReadOnlyList<PickLine> Lines);

/// <summary>
/// Загрузочники подтоварки. На каждое место хранения и каждый «комент» - свой лист:
/// «Очки отдельно» на «иза» и на «измп» - два разных загрузочника. Лист называется
/// «З‹место›-‹начало комента›»: «ЗМП-любой», «ЗА-мелкий», «ЗМПП-352».
///
/// У подразделения 206 загрузочники по местам не делятся - только по «коментам»:
/// «З-микс», «З-крупное».
/// </summary>
public static class RestockLoaderBuilder
{
    /// <summary>Самое длинное название листа, которое принимает Excel.</summary>
    public const int SheetNameLimit = 31;

    /// <summary>Подразделение, у которого загрузочники не делятся по местам хранения.</summary>
    public const string SinglePlaceDivision = "206";

    /// <summary>«Комент» заказа, у которого каждый код грузится своим номером заказа.</summary>
    public const string MonoMarker = "моно";

    public static IReadOnlyList<RestockLoader> Build(
        IReadOnlyList<PickLine> lines,
        Func<int, string> commentOfRow,
        bool mergePlaces = false)
    {
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<RestockLoader>();

        if (mergePlaces)
        {
            foreach (var group in lines.GroupBy(line => TextUtils.NormalizeKey(commentOfRow(line.RowIndex))))
            {
                var merged = TextUtils.Normalize(commentOfRow(group.First().RowIndex));
                result.Add(new RestockLoader(SheetName(null, merged, taken), null, merged, group.ToList()));
            }

            return result;
        }

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

    /// <summary>
    /// Номер заказа по строкам загрузочника. Обычно это номер города: Мск - 1, Спб - 2.
    /// У «моно» заказов каждый код грузится отдельно, и номер растёт по кодам.
    /// </summary>
    public static IReadOnlyList<double> OrderNumbers(RestockLoader loader, IReadOnlyList<string> cities)
    {
        if (TextUtils.ContainsKey(loader.Comment, MonoMarker))
        {
            var numbers = new Dictionary<string, double>(StringComparer.Ordinal);
            return loader.Lines
                .Select(line =>
                {
                    var key = TextUtils.NormalizeKey(line.Code);
                    if (!numbers.TryGetValue(key, out var number))
                    {
                        number = numbers.Count + 1;
                        numbers[key] = number;
                    }

                    return number;
                })
                .ToList();
        }

        return loader.Lines.Select(line => (double)Math.Max(CityNumber(cities, line.City), 1)).ToList();
    }

    private static int CityNumber(IReadOnlyList<string> cities, string city)
    {
        for (var index = 0; index < cities.Count; index++)
        {
            if (TextUtils.EqualsKey(cities[index], TextUtils.NormalizeKey(city)))
            {
                return index + 1;
            }
        }

        return 1;
    }

    /// <summary>«ЗМП-любой»: начало «комента» - первое слово или число, «352-148ЧЕРНЫЙ» даёт «352».</summary>
    public static string SheetName(StoragePlace? place, string comment, ISet<string> taken)
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

        var name = "З" + (place is { } value ? StoragePlaces.Code(value) : string.Empty) +
                   (start.Length > 0 ? "-" + start : string.Empty);
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
    /// «Комментарий» загрузочника: маркетплейс, город, «комент» заказа и приоритет.
    /// У поставок к этому добавляются их номера без повторов:
    /// «Lamoda Мск Подтоварка Любое из С318-156, С2437-027 Приоритет к 17.09».
    /// </summary>
    public static string CommentText(
        string marketplace,
        string city,
        string comment,
        IEnumerable<string> supplyNumbers,
        IEnumerable<string> priorities)
    {
        var parts = new List<string>();
        if (marketplace.Length > 0)
        {
            parts.Add(marketplace);
        }

        var cityText = TextUtils.Normalize(city);
        if (cityText.Length > 0)
        {
            parts.Add(cityText);
        }

        parts.Add("Подтоварка");

        var commentText = SentenceCase(comment);
        if (commentText.Length > 0)
        {
            parts.Add(commentText);
        }

        var supplies = string.Join(", ", Distinct(supplyNumbers));
        if (supplies.Length > 0)
        {
            parts.Add("из " + supplies);
        }

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

/// <summary>Какие строки «Р» остаются: заказы текущего года.</summary>
public static class ReservePeriod
{
    /// <summary>
    /// Прошлые годы удаляются: такие заказы уже отгружены или отменены. Внутри года
    /// остаётся всё - резерв из мая тоже держит товар.
    /// </summary>
    public static bool Keep(DateTime date, DateTime today) => date.Year >= today.Year;
}
