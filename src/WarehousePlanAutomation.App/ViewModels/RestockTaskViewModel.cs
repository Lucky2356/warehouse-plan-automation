using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>
/// Подтоварка маркетплейса: разбор колонки «Запрет» на листе «на загрузку»,
/// «Заметка» по каждой строке и лист согласования, затем места хранения,
/// листы «из‹место›» и загрузочники.
/// </summary>
public sealed class RestockTaskViewModel : WorkbookTaskViewModel
{
    public RestockTaskViewModel(
        IWorkbookProcessor processor, IAppLogger logger, UserSettings? settings = null)
        : base(processor, logger, settings)
    {
    }

    protected override string SettingsKey => "restock";

    protected override IReadOnlyList<string> RequiredSheets => new[]
    {
        RestockSchema.LoadSheet,
    };

    protected override IReadOnlyDictionary<string, string> OptionalSheets => new Dictionary<string, string>
    {
        [RestockSchema.AddressSheet] = "места хранения и загрузочники не соберутся",
        [RestockSchema.ReservesSheet] = "места хранения и загрузочники не соберутся",
        [RestockSchema.SuppliesSheet] = "места хранения и загрузочники не соберутся",
        [RestockSchema.ExceptionsSheet] =
            "исключений нет: всё, что с запретом забора из розницы, пойдёт на согласование",
    };

    public override string Title => "Подтоварка МП";

    public override string Subtitle => "«Заметка», места хранения и загрузочники";

    /// <summary>Коробка со стрелкой вниз: товар, который довозят на маркетплейс.</summary>
    public override string IconData =>
        "M3.2,7.4 L10,4.1 L16.8,7.4 L16.8,13.6 L10,16.9 L3.2,13.6 Z " +
        "M3.2,7.4 L10,10.7 L16.8,7.4 " +
        "M10,10.7 L10,16.9";

    public override string ActionCaption => "Разобрать подтоварку";

    public override string InitialHint =>
        "Выберите книгу подтоварки, в которой уже подтянуты «Запрет» и «Фактическое кол-во» " +
        "из сезонного файла и обновлены листы «по адресам и таре», «Р» и «МПП». Городов может быть " +
        "несколько: тогда нужна колонка «Итого в подтоварку». Что отдаём без согласования, " +
        "программа возьмёт с листа «Исключения», если он есть.";

    public override string SuccessMessage => "Подтоварка разобрана";

    public override string FailureMessage => "Не удалось разобрать подтоварку.";
}
