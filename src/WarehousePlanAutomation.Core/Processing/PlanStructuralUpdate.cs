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

/// <summary>
/// Заранее запланированная строка «Плана», в которую подставляется номер загрузки.
/// Аналитик заводит такие строки без номера, когда поставка ещё только планируется;
/// когда заказ появляется в выгрузке, строка не дублируется, а дополняется.
/// </summary>
public sealed record PlannedRowMatch(int ExcelRow, long LoadNumber, string Processing);

/// <summary>Изменение строки существующего или только что добавленного заказа.</summary>
/// <param name="Status">null означает «не менять существующее значение».</param>
/// <param name="CompletionPercent">null означает «не менять существующее значение».</param>
/// <param name="MissingFromOrders">
/// Заказ есть в «Плане», но сегодня пропал из «Заказов на отгрузку». Количество обнуляется,
/// а номер загрузки подсвечивается розовым, чтобы строку было видно и можно было убрать вручную.
/// </param>
/// <param name="IsNewRow">
/// Строка заведена сегодняшним запуском. Номер загрузки подсвечивается зелёным,
/// чтобы новые строки было видно среди вчерашних.
/// </param>
public sealed record OrderRowUpdate(
    long LoadNumber,
    double Quantity,
    string? Status,
    double? CompletionPercent,
    bool MissingFromOrders = false,
    bool IsNewRow = false);

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
        IReadOnlyList<AggregateUpdate> aggregateUpdates,
        IReadOnlyList<PlannedRowMatch>? plannedMatches = null)
    {
        PlanRowsToDelete = planRowsToDelete;
        NewRows = newRows;
        OrderUpdates = orderUpdates;
        AggregateUpdates = aggregateUpdates;
        PlannedMatches = plannedMatches ?? Array.Empty<PlannedRowMatch>();
    }

    /// <summary>Строки-заглушки «Заказы будут загружены», которые нужно удалить.</summary>
    public IReadOnlyList<int> PlanRowsToDelete { get; }

    public IReadOnlyList<NewPlanRowSpec> NewRows { get; }

    /// <summary>Строки без номера загрузки, в которые подставляется номер вместо вставки новой.</summary>
    public IReadOnlyList<PlannedRowMatch> PlannedMatches { get; }

    public IReadOnlyList<OrderRowUpdate> OrderUpdates { get; }

    public IReadOnlyList<AggregateUpdate> AggregateUpdates { get; }
}
