using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>
/// Приемка на хранилище, второй этап: «Допоставить» решено, программа собирает «адреса»
/// и «тары» и подтягивает на «итог» место хранения, тару и штрихкод.
/// </summary>
public sealed class ReceivingAddressesTaskViewModel : WorkbookTaskViewModel
{
    public ReceivingAddressesTaskViewModel(
        IWorkbookProcessor processor, IAppLogger logger, UserSettings? settings = null)
        : base(processor, logger, settings)
    {
    }

    protected override string SettingsKey => "receivingAddresses";

    protected override IReadOnlyList<string> RequiredSheets => new[]
    {
        ReceivingSchema.SummarySheet,
        ReceivingSchema.StorageSheet,
        ReceivingSchema.MarketplaceSheet,
        ReceivingSchema.CollectedSheet,
        ReceivingSchema.AddressesSheet,
        ReceivingSchema.ContainersSheet,
        ReceivingSchema.MarketplaceAddressesSheet,
        ReceivingSchema.MarketplaceContainersSheet,
    };

    /// <summary>Именно «Приемка на хранилище»: просто «Приемка» - другой отчёт.</summary>
    public override string Title => "Приемка на хранилище · адреса";

    public override string Subtitle => "Приемка на хранилище: «Допоставить» заполнено → адреса, тары, место хранения";

    /// <summary>Метка на карте: у товара появляется адрес.</summary>
    public override string IconData =>
        "M10,17.4 C7.4,14.4 4.8,11.2 4.8,8.1 A5.2,5.2 0 0,1 15.2,8.1 C15.2,11.2 12.6,14.4 10,17.4 Z " +
        "M11.9,8.1 A1.9,1.9 0 1,1 8.1,8.1 A1.9,1.9 0 1,1 11.9,8.1";

    public override string ActionCaption => "Расставить адреса";

    public override string InitialHint =>
        "Выберите файл приемки на хранилище после подготовки - тот, в котором на листе «итог» уже решено «Допоставить».";

    public override string SuccessMessage => "Адреса расставлены";

    public override string FailureMessage => "Не удалось расставить адреса.";
}
