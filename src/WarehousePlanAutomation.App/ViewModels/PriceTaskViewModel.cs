using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>
/// Первый этап распреда: из инвойса и справочников собираются листы «для цен» и «Цены».
/// После него аналитик протягивает колонки со стенками - и запускает второй этап.
/// </summary>
public sealed class PriceTaskViewModel : WorkbookTaskViewModel
{
    public PriceTaskViewModel(
        IWorkbookProcessor processor, IAppLogger logger, UserSettings? settings = null)
        : base(processor, logger, settings)
    {
    }

    protected override string SettingsKey => "price";

    protected override IReadOnlyList<string> RequiredSheets => new[]
    {
        PriceSchema.InvoiceSheet,
        PriceSchema.ForPricesSheet,
        PriceSchema.PricesSheet,
        PriceSchema.SummaryPriceSheet,
        PriceSchema.DivisionPriceSheet,
    };

    /// <summary>Инвойс - лист на каждую поставку: «Invoice-338», «Invoice-341».</summary>
    protected override IReadOnlyCollection<string> RepeatableSheets { get; } = new[] { PriceSchema.InvoiceSheet };

    public override string Title => "Распред · подготовка";

    public override string Subtitle => "Инвойс и справочники → «для цен» и «Цены»";

    /// <summary>Развилка: одна поставка расходится по магазинам.</summary>
    public override string IconData =>
        "M11.5,2.6 A1.5,1.5 0 1,1 8.5,2.6 A1.5,1.5 0 1,1 11.5,2.6 " +
        "M10,4.1 L10,7.8 " +
        "M4.4,17.4 L4.4,13.3 A1.6,1.6 0 0,1 6,11.7 L14,11.7 A1.6,1.6 0 0,1 15.6,13.3 L15.6,17.4 " +
        "M10,11.7 L10,17.4";

    public override string ActionCaption => "Подготовить распред";

    public override string InitialHint =>
        "Выберите книгу распреда с обновлёнными листами «Invoice» (по листу на поставку), " +
        "«Сводный прайс» и «Прайс по подразделениям». Сектор берётся из названия файла: " +
        "один распред - один сектор.";

    public override string SuccessMessage => "Данные для распреда подготовлены";

    public override string FailureMessage => "Не удалось подготовить распред.";
}
