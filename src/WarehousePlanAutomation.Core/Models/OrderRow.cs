namespace WarehousePlanAutomation.Core.Models;

/// <summary>Строка листа «Заказы на отгрузку».</summary>
public sealed record OrderRow(
    int ExcelRow,
    string Division,
    string Comment,
    double DifferenceUnits,
    double? DocumentDate);

/// <summary>Категория, к которой отнесена строка выгрузки.</summary>
public enum OrderCategory
{
    /// <summary>Подразделение содержит «Опт».</summary>
    Wholesale,

    /// <summary>Подразделение содержит «777».</summary>
    InternetShop,

    /// <summary>Маркетплейс, комментарий содержит «возвр».</summary>
    MarketplaceReturns,

    /// <summary>Маркетплейс, в комментарии найден номер поставки.</summary>
    MarketplaceSupplies,

    /// <summary>Оставшиеся строки маркетплейсов.</summary>
    MarketplaceStorage,

    /// <summary>Служебные строки: «автозаказ», «вирт», «фото».</summary>
    Service,

    /// <summary>Реальный заказ с номером загрузки.</summary>
    RealOrder,

    /// <summary>Строка не подошла ни под одно правило и остаётся для ручной проверки.</summary>
    Unresolved,
}

/// <summary>Сегодняшний заказ, сгруппированный по номеру загрузки.</summary>
public sealed record TodayOrderGroup(
    long LoadNumber,
    double Quantity,
    string Comment,
    double? DocumentDate,
    IReadOnlyList<int> SourceRows);

/// <summary>Результат классификации листа «Заказы на отгрузку».</summary>
public sealed class OrderClassificationResult
{
    public OrderClassificationResult(
        IReadOnlyDictionary<OrderCategory, double> totals,
        IReadOnlyList<int> rowsToDelete,
        IReadOnlyList<TodayOrderGroup> groups,
        IReadOnlyList<OrderRow> leftovers)
    {
        Totals = totals;
        RowsToDelete = rowsToDelete;
        Groups = groups;
        Leftovers = leftovers;
    }

    /// <summary>Суммы «разница единиц» по категориям. Отсутствующая категория означает 0.</summary>
    public IReadOnlyDictionary<OrderCategory, double> Totals { get; }

    /// <summary>Строки Excel, которые должны быть физически удалены с листа.</summary>
    public IReadOnlyList<int> RowsToDelete { get; }

    /// <summary>Сегодняшние реальные заказы, сгруппированные по номеру загрузки.</summary>
    public IReadOnlyList<TodayOrderGroup> Groups { get; }

    /// <summary>Строки, оставленные на листе для ручного контроля.</summary>
    public IReadOnlyList<OrderRow> Leftovers { get; }

    public double Total(OrderCategory category) => Totals.TryGetValue(category, out var value) ? value : 0d;
}
