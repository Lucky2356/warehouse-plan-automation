using WarehousePlanAutomation.Core.Models;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Агрегатные строки листа «План», значение которых задаётся суммами выгрузки.</summary>
public enum PlanAggregateTarget
{
    AutoHub,
    Wholesale,
    InternetShop,
    MarketplaceFromStorage,
    MarketplaceFromReturns,
    MarketplaceFromSupplies,
}

/// <summary>Новая строка заказа, которую нужно добавить в блок листа «План».</summary>
public sealed record NewPlanRowSpec(
    PlanSectionKind Section,
    string Supplies,
    string Processing,
    string Comments,
    double? DocumentDate,
    long LoadNumber);

/// <summary>Изменение строки существующего или только что добавленного заказа.</summary>
/// <param name="Status">null означает «не менять существующее значение».</param>
/// <param name="CompletionPercent">null означает «не менять существующее значение».</param>
public sealed record OrderRowUpdate(long LoadNumber, double Quantity, string? Status, double? CompletionPercent);

/// <summary>Изменение агрегатной строки.</summary>
public sealed record AggregateUpdate(PlanAggregateTarget Target, double Quantity);

/// <summary>
/// Полный набор изменений листа «План», рассчитанный в памяти до обращения к Excel.
/// Порядок применения: удаление строк, вставка новых, обновление значений.
/// </summary>
public sealed class PlanStructuralUpdate
{
    public PlanStructuralUpdate(
        IReadOnlyList<int> planRowsToDelete,
        IReadOnlyList<NewPlanRowSpec> newRows,
        IReadOnlyList<OrderRowUpdate> orderUpdates,
        IReadOnlyList<AggregateUpdate> aggregateUpdates)
    {
        PlanRowsToDelete = planRowsToDelete;
        NewRows = newRows;
        OrderUpdates = orderUpdates;
        AggregateUpdates = aggregateUpdates;
    }

    /// <summary>Строки-заглушки «Заказы будут загружены», которые нужно удалить.</summary>
    public IReadOnlyList<int> PlanRowsToDelete { get; }

    public IReadOnlyList<NewPlanRowSpec> NewRows { get; }

    public IReadOnlyList<OrderRowUpdate> OrderUpdates { get; }

    public IReadOnlyList<AggregateUpdate> AggregateUpdates { get; }
}
