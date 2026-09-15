using System.Globalization;
using System.Text.RegularExpressions;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Место хранения, с которого собирается подтоварка.</summary>
public enum StoragePlace
{
    /// <summary>«МП» - адреса маркетплейса.</summary>
    Marketplace,

    /// <summary>«МПП» - непринятый товар в полученных поставках с номерами на «М» и «Л».</summary>
    Supplies,

    /// <summary>«А» - хранение и хранилище.</summary>
    Storage,

    /// <summary>«СЗП» - непринятый товар в поставках с номерами на «С».</summary>
    NetworkSupplies,

    /// <summary>«В» - возвраты, времянка, олды.</summary>
    Returns,
}

public static class StoragePlaces
{
    /// <summary>
    /// Порядок, в котором берём товар. Он же зашит в названиях колонок «на загрузку»:
    /// «1МП», «2МПП», «3А», «4СЗП», «5В».
    /// </summary>
    public static readonly IReadOnlyList<StoragePlace> Priority = new[]
    {
        StoragePlace.Marketplace,
        StoragePlace.Supplies,
        StoragePlace.Storage,
        StoragePlace.NetworkSupplies,
        StoragePlace.Returns,
    };

    /// <summary>Как место называется на листах и в тексте «Места хранения».</summary>
    public static string Code(StoragePlace place) => place switch
    {
        StoragePlace.Marketplace => "МП",
        StoragePlace.Supplies => "МПП",
        StoragePlace.Storage => "А",
        StoragePlace.NetworkSupplies => "СЗП",
        _ => "В",
    };

    /// <summary>Колонка «на загрузку» с остатком места: номер - его очередь.</summary>
    public static string LoadColumn(StoragePlace place) =>
        (Priority.ToList().IndexOf(place) + 1).ToString(CultureInfo.InvariantCulture) + Code(place);

    /// <summary>Лист, куда выносятся строки этого места: «измп», «иза», «измпп».</summary>
    public static string PickSheet(StoragePlace place) => "из" + Code(place).ToLowerInvariant();

    /// <summary>Остаток поставок - это «Отклонение», у адресов - «Количество».</summary>
    public static bool IsSupply(StoragePlace place) =>
        place is StoragePlace.Supplies or StoragePlace.NetworkSupplies;

    /// <summary>
    /// Лист «по адресам и таре»: «Маркетплейс» - на «МП», «Хранение» и «Хранилище» - на «А»,
    /// «Возвраты», «Времянка» и «Олды» - на «В». В выгрузке олды пишутся «ОЛД».
    /// Остальное («Образцы», «Брак уценка», «Нет Маркировки») в подтоварку не идёт.
    /// </summary>
    public static StoragePlace? FromStorageType(string? storageType) => TextUtils.NormalizeKey(storageType) switch
    {
        "маркетплейс" => StoragePlace.Marketplace,
        "хранение" or "хранилище" => StoragePlace.Storage,
        "возвраты" or "возврат" or "времянка" or "олд" or "олды" => StoragePlace.Returns,
        _ => null,
    };

    /// <summary>Номер непринятой поставки: «С…» - на «СЗП», «М…» и «Л…» - на «МПП».</summary>
    public static StoragePlace? FromSupplyNumber(string? number)
    {
        var letter = SupplyNumbers.Letter(number);
        return letter switch
        {
            SupplyNumbers.Network => StoragePlace.NetworkSupplies,
            SupplyNumbers.Marketplace or SupplyNumbers.Excluded => StoragePlace.Supplies,
            _ => null,
        };
    }
}

/// <summary>Что делать с резервом при расстановке места хранения.</summary>
public enum ReserveKind
{
    /// <summary>Резерв лежит на определённом месте, его вычитаем оттуда.</summary>
    Place,

    /// <summary>«Опт» - может взять с любого места, конкретному месту резерв не мешает.</summary>
    Anywhere,

    /// <summary>«На образцы» - резервы не учитываем.</summary>
    Samples,

    /// <summary>По комментарию место не понять.</summary>
    Unknown,
}

public readonly record struct ReserveTarget(ReserveKind Kind, StoragePlace? Place);

/// <summary>Резерв с листа «Р»: сколько и под какой заказ.</summary>
public sealed record ReserveLine(double Quantity, string Comment);

/// <summary>
/// Откуда резерв, понятно по комментарию заказа: «Золотое Яблоко Сумки из адресов МП,
/// А1,А2,А3, Возвратов приоритет к 04.09» - резерв на МП, место названо первым;
/// «Срочная подтоварка 09.09_Хранение, хранилище» - на «А»; «Заказ интернет магазина» -
/// тоже «А»: интернет-магазин грузится с хранения.
/// </summary>
public static class ReserveTargets
{
    private static readonly Regex Wholesale = new(@"(?<![\p{L}])опт", RegexOptions.CultureInvariant);

    private static readonly (Regex Pattern, StoragePlace Place)[] Places =
    {
        (new Regex(@"адрес\w*\s+мп(?![\p{L}])", RegexOptions.CultureInvariant), StoragePlace.Marketplace),
        (new Regex(@"(?<![\p{L}\d])а[123](?!\d)", RegexOptions.CultureInvariant), StoragePlace.Storage),
        (new Regex(@"хранени|хранилищ|интер\w*\s*-?\s*магазин", RegexOptions.CultureInvariant), StoragePlace.Storage),
        (new Regex(@"возврат|времянк", RegexOptions.CultureInvariant), StoragePlace.Returns),
        (new Regex(@"(?<![\p{L}\d])[млm]з?\s?\d{2,5}-\d{2,4}", RegexOptions.CultureInvariant), StoragePlace.Supplies),
        (new Regex(@"(?<![\p{L}\d])[сc]з?\s?\d{2,5}-\d{2,4}", RegexOptions.CultureInvariant), StoragePlace.NetworkSupplies),
    };

    public static ReserveTarget Parse(string? comment)
    {
        var key = TextUtils.NormalizeKey(comment);

        if (key.Contains("образц", StringComparison.Ordinal))
        {
            return new ReserveTarget(ReserveKind.Samples, null);
        }

        if (Wholesale.IsMatch(key))
        {
            return new ReserveTarget(ReserveKind.Anywhere, null);
        }

        StoragePlace? found = null;
        var position = int.MaxValue;
        foreach (var (pattern, place) in Places)
        {
            var match = pattern.Match(key);
            if (match.Success && match.Index < position)
            {
                position = match.Index;
                found = place;
            }
        }

        return found is null
            ? new ReserveTarget(ReserveKind.Unknown, null)
            : new ReserveTarget(ReserveKind.Place, found);
    }
}

/// <summary>Сколько берём с одного места.</summary>
public sealed record AllocationPart(StoragePlace Place, double Quantity);

/// <summary>Решение по строке «на загрузку»: откуда берём и хватило ли.</summary>
/// <param name="Parts">Места и количества по очереди приоритета. Пусто - взять неоткуда.</param>
/// <param name="Missing">Сколько не хватило: больше нуля - строка «Не собрано».</param>
/// <param name="UnknownReserves">
/// Резервы, место которых по комментарию не понять. Они ни из чего не вычтены - количество
/// могло оказаться завышенным.
/// </param>
public sealed record StorageAllocation(
    IReadOnlyList<AllocationPart> Parts,
    double Missing,
    IReadOnlyList<ReserveLine> UnknownReserves)
{
    public bool Complete => Missing <= 0d;

    /// <summary>Всё количество с одного места: в «Месте хранения» просто его название.</summary>
    public bool SinglePlace => Complete && Parts.Count == 1;

    public double Collected => Parts.Sum(part => part.Quantity);

    /// <summary>
    /// «МП» - всё с одного места; «МП10, А5» - собрано с нескольких; «А2» или «0» - не хватило,
    /// цифра говорит, сколько на самом деле есть.
    /// </summary>
    public string Text
    {
        get
        {
            if (SinglePlace)
            {
                return StoragePlaces.Code(Parts[0].Place);
            }

            return Parts.Count == 0
                ? "0"
                : string.Join(", ", Parts.Select(part => StoragePlaces.Code(part.Place) + Quantity(part.Quantity)));
        }
    }

    public static string Quantity(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);
}

/// <summary>
/// Место хранения строки «на загрузку».
///
/// Сначала ищется одно место, с которого можно взять всё: остаток места минус все резервы АЦР
/// минус «в подтоварку» не меньше нуля. Места перебираются в порядке МП, МПП, А, СЗП, В.
///
/// Если такого нет, количество набирается с нескольких мест по той же очереди. Резервы при этом
/// вычитаются только из того места, на котором лежат, - это видно по комментарию заказа.
/// «Опт» может взять откуда угодно, а на образцы резервы не учитываются вовсе.
/// </summary>
public static class StorageAllocator
{
    public static StorageAllocation Allocate(
        double need,
        IReadOnlyDictionary<StoragePlace, double> stock,
        IReadOnlyList<ReserveLine> reserves)
    {
        var totalReserve = reserves.Sum(reserve => reserve.Quantity);

        foreach (var place in StoragePlaces.Priority)
        {
            if (Stock(stock, place) - totalReserve - need >= 0d)
            {
                return new StorageAllocation(
                    new[] { new AllocationPart(place, need) }, 0d, Array.Empty<ReserveLine>());
            }
        }

        var available = StoragePlaces.Priority.ToDictionary(place => place, place => Math.Max(Stock(stock, place), 0d));
        var unknown = new List<ReserveLine>();

        foreach (var reserve in reserves)
        {
            var target = ReserveTargets.Parse(reserve.Comment);
            switch (target.Kind)
            {
                case ReserveKind.Place:
                    var place = target.Place!.Value;
                    available[place] = Math.Max(available[place] - reserve.Quantity, 0d);
                    break;

                case ReserveKind.Unknown:
                    unknown.Add(reserve);
                    break;
            }
        }

        var parts = new List<AllocationPart>();
        var remaining = need;
        foreach (var place in StoragePlaces.Priority)
        {
            if (remaining <= 0d)
            {
                break;
            }

            var take = Math.Min(available[place], remaining);
            if (take > 0d)
            {
                parts.Add(new AllocationPart(place, take));
                remaining -= take;
            }
        }

        return new StorageAllocation(parts, Math.Max(remaining, 0d), unknown);
    }

    private static double Stock(IReadOnlyDictionary<StoragePlace, double> stock, StoragePlace place) =>
        stock.TryGetValue(place, out var value) ? value : 0d;
}
