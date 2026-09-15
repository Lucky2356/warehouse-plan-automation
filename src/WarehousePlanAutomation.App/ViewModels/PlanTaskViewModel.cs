using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>Формирование листа «План» из вчерашнего плана и двух свежих выгрузок.</summary>
public sealed class PlanTaskViewModel : WorkbookTaskViewModel
{
    public PlanTaskViewModel(
        IWorkbookProcessor processor, IAppLogger logger, UserSettings? settings = null)
        : base(processor, logger, settings)
    {
    }

    protected override string SettingsKey => "plan";

    protected override IReadOnlyList<string> RequiredSheets => new[]
    {
        SheetSchema.PlanSheet,
        SheetSchema.OrdersSheet,
        SheetSchema.JournalSheet,
    };

    public override string Title => "План склада";

    public override string Subtitle => "Вчерашний план плюс свежие выгрузки";

    /// <summary>Стопка слоёв: план собирается из нескольких листов.</summary>
    public override string IconData =>
        "M3,6.2 L10,3 L17,6.2 L10,9.4 Z " +
        "M3,10.2 L10,13.4 L17,10.2 " +
        "M3,13.9 L10,17 L17,13.9";

    public override string ActionCaption => "Сформировать план склада";

    public override string InitialHint =>
        "Выберите вчерашнюю книгу с обновлёнными листами «Заказы на отгрузку» " +
        "и «Журнал заказов на отгрузку».";

    public override string SuccessMessage => "План склада успешно сформирован";

    public override string FailureMessage => "Не удалось сформировать план склада.";
}
