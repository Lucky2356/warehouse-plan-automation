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

    /// <summary>«В» - возвраты и времянка.</summary>
    Returns,
}

public static class StoragePlaces
{
    /// <summary>
    /// Порядок, в котором берём товар. Он же зашит в названиях колонок «на загрузку»:
    /// «1МП», «2МПП», «3В», «4А», «5СЗП».
    /// </summary>
    public static readonly IReadOnlyList<StoragePlace> Priority = new[]
    {
        StoragePlace.Marketplace,
        StoragePlace.Supplies,
        StoragePlace.Returns,
        StoragePlace.Storage,
        StoragePlace.NetworkSupplies,
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
    /// «Возвраты» и «Времянка» - на «В». Остальное - олды («ОЛД» в выгрузке), «Образцы»,
    /// «Брак уценка», «Нет Маркировки» - в подтоварку не идёт.
    /// </summary>
    public static StoragePlace? FromStorageType(string? storageType) => TextUtils.NormalizeKey(storageType) switch
    {
        "маркетплейс" => StoragePlace.Marketplace,
        "хранение" or "хранилище" => StoragePlace.Storage,
        "возвраты" or "возврат" or "времянка" => StoragePlace.Returns,
        _ => null,
    };

    /// <summary>
    /// Загрузочники без разбивки по местам: один на адреса (МП, А, В), другой на поставки
    /// (МПП, СЗП). Порядок - тот, в котором строка ищет себе загрузочник.
    /// </summary>
    public static readonly IReadOnlyList<IReadOnlyList<StoragePlace>> Groups = new[]
    {
        (IReadOnlyList<StoragePlace>)new[] { StoragePlace.Marketplace, StoragePlace.Returns, StoragePlace.Storage },
        new[] { StoragePlace.Supplies, StoragePlace.NetworkSupplies },
    };

    /// <summary>Загрузочник места при работе без разбивки: 0 - адреса, 1 - поставки.</summary>
    public static int GroupOf(StoragePlace place) => IsSupply(place) ? 1 : 0;

    /// <summary>Как группа называется в названии загрузочника: «МП+А+В», «МПП+СЗП».</summary>
    public static string GroupCode(int group) => group == 0 ? "МП+А+В" : "МПП+СЗП";

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

    /// <summary>«Опт» - может взять с любого места, где есть остаток.</summary>
    Anywhere,

    /// <summary>«На образцы» и «на фото» - берут только с поставок: МПП и СЗП.</summary>
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

        if (key.Contains("образц", StringComparison.Ordinal) || key.Contains("на фото", StringComparison.Ordinal))
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
/// Резервы, место которых по комментарию не понять. Они вычтены с того места, где был
/// остаток, - комментарий стоит проверить.
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
/// Сначала ищется одно место, с которого можно взять всё. Места перебираются в порядке
/// МП, МПП, В, А, СЗП, и резервы АЦР для проверки места уводятся на другие места: если
/// на МП хватает на подтоварку, а резервы покрывает СЗП, берём с МП. На самом месте резерв
/// остаётся, только когда он лежит именно там (так по комментарию заказа) или другим местам
/// его не покрыть.
///
/// Если такого нет, количество набирается с нескольких мест по той же очереди. Резервы при этом
/// вычитаются из того места, на котором лежат, - это видно по комментарию заказа. Резервы опта
/// снимаются с любого места, где есть остаток, а на образцы и на фото - только с поставок,
/// то есть с МПП и СЗП. Резерв, место которого по комментарию не понять, тоже вычитается -
/// с любого места, как опт: товар под ним уже занят. О нём только пишется замечание.
///
/// У строк «Отгрузка в рамках заказа МП» место резерва по комментарию не важно: подтоварка
/// в первую очередь берётся с МП и МПП, а резервы уводятся на другие места - В, А, СЗП;
/// на МП и МПП они ложатся, только если остальным местам их не покрыть.
/// </summary>
public static class StorageAllocator
{
    /// <summary>Места маркетплейса: с них можно брать и при запрете забора из розницы.</summary>
    public static readonly IReadOnlyList<StoragePlace> MarketplacePlaces = new[]
    {
        StoragePlace.Marketplace,
        StoragePlace.Supplies,
    };

    /// <param name="only">
    /// Брать только с этих мест. null - со всех. Резервы при этом по-прежнему могут лечь
    /// на любое место.
    /// </param>
    public static StorageAllocation Allocate(
        double need,
        IReadOnlyDictionary<StoragePlace, double> stock,
        IReadOnlyList<ReserveLine> reserves,
        bool marketplaceFirst = false,
        IReadOnlyCollection<StoragePlace>? only = null)
    {
        var places = StoragePlaces.Priority.Where(place => only is null || only.Contains(place)).ToList();
        var unknown = reserves
            .Where(reserve => ReserveTargets.Parse(reserve.Comment).Kind == ReserveKind.Unknown)
            .ToList();

        foreach (var place in places)
        {
            // Всё с одного места: непонятные резервы легли на другие места, замечание о них
            // только засорило бы список.
            if (Residual(stock, reserves, marketplaceFirst, new[] { place })[place] >= need)
            {
                return new StorageAllocation(
                    new[] { new AllocationPart(place, need) }, 0d, Array.Empty<ReserveLine>());
            }
        }

        var available = only is null
            ? Split(stock, reserves, marketplaceFirst)
            : Residual(stock, reserves, marketplaceFirst, only);

        var parts = new List<AllocationPart>();
        var remaining = need;
        foreach (var place in places)
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

    /// <summary>
    /// Расстановка без разбивки по местам: строка целиком уходит в загрузочник адресов,
    /// если МП + А + В за вычетом всех резервов АЦР покрывают «в подтоварку», иначе -
    /// в загрузочник поставок, если его покрывают МПП + СЗП. Не хватает ни там, ни там -
    /// количество набирается как обычно, и части попадут в оба загрузочника.
    /// </summary>
    public static StorageAllocation AllocateByGroups(
        double need,
        IReadOnlyDictionary<StoragePlace, double> stock,
        IReadOnlyList<ReserveLine> reserves,
        bool marketplaceFirst = false)
    {
        var totalReserve = reserves.Sum(reserve => reserve.Quantity);
        foreach (var group in StoragePlaces.Groups)
        {
            if (group.Sum(place => Math.Max(Stock(stock, place), 0d)) - totalReserve >= need)
            {
                return Allocate(need, stock, reserves, marketplaceFirst, group);
            }
        }

        return Allocate(need, stock, reserves, marketplaceFirst);
    }

    /// <summary>
    /// Хватает ли одних МП и МПП, если резервы увести на другие места: тогда строке
    /// с запретом забора из розницы согласование не нужно.
    /// </summary>
    public static bool MarketplaceCovers(
        double need,
        IReadOnlyDictionary<StoragePlace, double> stock,
        IReadOnlyList<ReserveLine> reserves)
    {
        var available = Residual(stock, reserves, marketplaceFirst: true, MarketplacePlaces);
        return MarketplacePlaces.Sum(place => available[place]) >= need;
    }

    /// <summary>
    /// Остатки мест за вычетом резервов, которые по возможности уведены с мест <paramref name="keep"/>.
    /// Сначала ложатся резервы с известным местом, потом на образцы (только МПП и СЗП),
    /// последними - те, что могут лечь куда угодно: так местам <paramref name="keep"/>
    /// остаётся больше всего.
    /// </summary>
    private static Dictionary<StoragePlace, double> Residual(
        IReadOnlyDictionary<StoragePlace, double> stock,
        IReadOnlyList<ReserveLine> reserves,
        bool marketplaceFirst,
        IReadOnlyCollection<StoragePlace> keep)
    {
        var available = StoragePlaces.Priority.ToDictionary(place => place, place => Math.Max(Stock(stock, place), 0d));
        var anywhere = LastOf(marketplaceFirst ? ReservePlacesForOrder : StoragePlaces.Priority, keep);
        var samples = LastOf(SamplePlaces, keep);

        var targets = reserves.Select(reserve => (reserve.Quantity, Target: ReserveTargets.Parse(reserve.Comment))).ToList();

        if (!marketplaceFirst)
        {
            foreach (var (quantity, target) in targets.Where(item => item.Target.Kind == ReserveKind.Place))
            {
                var place = target.Place!.Value;
                available[place] = Math.Max(available[place] - quantity, 0d);
            }
        }

        foreach (var (quantity, _) in targets.Where(item => item.Target.Kind == ReserveKind.Samples))
        {
            Subtract(available, samples, quantity);
        }

        foreach (var (quantity, target) in targets)
        {
            var flexible = target.Kind is ReserveKind.Anywhere or ReserveKind.Unknown ||
                           (marketplaceFirst && target.Kind == ReserveKind.Place);
            if (flexible)
            {
                Subtract(available, anywhere, quantity);
            }
        }

        return available;
    }

    /// <summary>Те же места, но <paramref name="keep"/> - в самом конце очереди.</summary>
    private static IReadOnlyList<StoragePlace> LastOf(IReadOnlyList<StoragePlace> places, IReadOnlyCollection<StoragePlace> keep) =>
        places.Where(place => !keep.Contains(place)).Concat(places.Where(keep.Contains)).ToList();

    /// <summary>Остатки для набора с нескольких мест: резервы - в порядке листа «Р».</summary>
    private static Dictionary<StoragePlace, double> Split(
        IReadOnlyDictionary<StoragePlace, double> stock,
        IReadOnlyList<ReserveLine> reserves,
        bool marketplaceFirst)
    {
        var available = StoragePlaces.Priority.ToDictionary(place => place, place => Math.Max(Stock(stock, place), 0d));

        foreach (var reserve in reserves)
        {
            var target = ReserveTargets.Parse(reserve.Comment);

            // На образцы и на фото берут из поставок - это правило сильнее любого другого.
            if (target.Kind == ReserveKind.Samples)
            {
                Subtract(available, SamplePlaces, reserve.Quantity);
                continue;
            }

            if (marketplaceFirst)
            {
                Subtract(available, ReservePlacesForOrder, reserve.Quantity);
                continue;
            }

            // Резерв снимается с мест по очереди: пока он не закончится или места не опустеют.
            switch (target.Kind)
            {
                case ReserveKind.Place:
                    var place = target.Place!.Value;
                    available[place] = Math.Max(available[place] - reserve.Quantity, 0d);
                    break;

                default:
                    Subtract(available, StoragePlaces.Priority, reserve.Quantity);
                    break;
            }
        }

        return available;
    }

    /// <summary>Откуда берут резервы на образцы и на фото.</summary>
    private static readonly IReadOnlyList<StoragePlace> SamplePlaces = new[]
    {
        StoragePlace.Supplies,
        StoragePlace.NetworkSupplies,
    };

    /// <summary>
    /// Откуда снимаются резервы у строки «Отгрузка в рамках заказа МП»: сначала места,
    /// которые подтоварке не нужны, и только потом МП и МПП.
    /// </summary>
    private static readonly IReadOnlyList<StoragePlace> ReservePlacesForOrder = new[]
    {
        StoragePlace.Returns,
        StoragePlace.Storage,
        StoragePlace.NetworkSupplies,
        StoragePlace.Marketplace,
        StoragePlace.Supplies,
    };

    private static void Subtract(
        IDictionary<StoragePlace, double> available, IReadOnlyList<StoragePlace> places, double quantity)
    {
        var remaining = quantity;
        foreach (var place in places)
        {
            if (remaining <= 0d)
            {
                return;
            }

            var take = Math.Min(available[place], remaining);
            available[place] -= take;
            remaining -= take;
        }
    }

    private static double Stock(IReadOnlyDictionary<StoragePlace, double> stock, StoragePlace place) =>
        stock.TryGetValue(place, out var value) ? value : 0d;
}
