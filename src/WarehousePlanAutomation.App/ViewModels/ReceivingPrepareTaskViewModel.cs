using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>
/// Приемка на хранилище, первый этап: выгрузки разобраны, «Приходы» покрашены,
/// листы поставок и «итог» собраны. После него аналитик решает «Допоставить».
/// </summary>
public sealed class ReceivingPrepareTaskViewModel : WorkbookTaskViewModel
{
    public ReceivingPrepareTaskViewModel(
        IWorkbookProcessor processor, IAppLogger logger, UserSettings? settings = null)
        : base(processor, logger, settings)
    {
    }

    protected override string SettingsKey => "receiving";

    /// <summary>«Согласование количества» сюда не входит: по инструкции его может и не быть.</summary>
    protected override IReadOnlyList<string> RequiredSheets => new[]
    {
        ReceivingSchema.StockSheet,
        ReceivingSchema.StorageStockSheet,
        ReceivingSchema.ReservesSheet,
        ReceivingSchema.StorageSheet,
        ReceivingSchema.MarketplaceSheet,
        ReceivingSchema.IncomingSheet,
        ReceivingSchema.SummarySheet,
        ReceivingSchema.CollectedSheet,
        ReceivingSchema.NotCollectedSheet,
        ReceivingSchema.MarketplaceSuppliesSheet,
        ReceivingSchema.WarehouseSheet,
        ReceivingSchema.RemovedFromPlanSheet,
        ReceivingSchema.PlanSheet,
        ReceivingSchema.NotAcceptedSheet,
    };

    /// <summary>
    /// Именно «Приемка на хранилище»: просто «Приемка» - другой отчёт, и путать их нельзя.
    /// </summary>
    public override string Title => "Приемка на хранилище · подготовка";

    public override string Subtitle => "Приемка на хранилище: выгрузки → «Приходы», поставки, «итог»";

    /// <summary>Стеллаж с полками: товар, который принимают на хранение.</summary>
    public override string IconData =>
        "M3.5,17 L3.5,3.5 L16.5,3.5 L16.5,17 " +
        "M3.5,8 L16.5,8 " +
        "M3.5,12.5 L16.5,12.5 " +
        "M6.5,5.8 L9.5,5.8 " +
        "M6.5,10.3 L11.5,10.3 " +
        "M6.5,14.8 L9,14.8";

    public override string ActionCaption => "Подготовить приемку на хранилище";

    public override string InitialHint =>
        "Выберите книгу «Приемка на хранилище», в которой уже обновлены «План склада», «Удалили из плана склада», " +
        "«Непринятый товар», «Согласование количества», «Склад», «Т.Остатки», «Остатки», «Резервы» " +
        "и «Приходы».";

    public override string SuccessMessage => "Приемка на хранилище подготовлена - заполните «Допоставить» на «итоге»";

    public override string FailureMessage => "Не удалось подготовить приемку на хранилище.";
}
