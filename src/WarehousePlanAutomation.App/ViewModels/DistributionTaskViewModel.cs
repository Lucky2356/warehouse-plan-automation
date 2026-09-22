using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Excel;

namespace WarehousePlanAutomation.App.ViewModels;

/// <summary>
/// Второй этап распреда: колонки со стенками уже протянуты, программа ставит наценки,
/// аналоги и согласованные цены, проверяет строки и заполняет листы «остатки», «Распред»
/// и «Загрузочник». Инвойс на этом этапе не читается - всё берётся с листа «Цены».
/// </summary>
public sealed class DistributionTaskViewModel : WorkbookTaskViewModel
{
    /// <summary>Под этим ключом помнится папка согласования цен. Её же читает обработчик при запуске.</summary>
    public const string ApprovedPricesFolderKey = "approvedPricesFolder";

    public DistributionTaskViewModel(
        IWorkbookProcessor processor, IAppLogger logger, UserSettings? settings = null)
        : base(processor, logger, settings)
    {
    }

    protected override string SettingsKey => "distribution";

    protected override IReadOnlyList<string> RequiredSheets => new[]
    {
        PriceSchema.PricesSheet,
        PriceSchema.MarkupSheet,
        PriceSchema.LinkSheet,
        PriceSchema.StockSourceSheet,
        PriceSchema.StockSheet,
        PriceSchema.DistributionSheet,
        PriceSchema.LoaderSheet,
        PriceSchema.SeasonalitySheet,
    };

    protected override IReadOnlyDictionary<string, string> OptionalSheets { get; } = new Dictionary<string, string>
    {
        [PriceSchema.AnaloguesSheet] = "«Аналог гугл» не заполнится и сверится таким, как стоит в книге",
    };

    protected override string? FolderOptionKey => ApprovedPricesFolderKey;

    public override string FolderOptionTitle => "СОГЛАСОВАНИЕ ЦЕН: ФАЙЛ ИЛИ ПАПКА";

    public override bool FolderOptionAllowsFile => true;

    public override string FolderOptionHint =>
        "Необязательно. Файл - цены берутся из него, а недостающие ищутся в его папке. Папка - программа " +
        "сама найдёт в ней файл по номеру поставки (при нехватке - ближайший по номеру); подойдёт папка " +
        "сезона или вся «Согласование цен» у КМ. Ничего не выбрано - «Согласованные цены» берутся " +
        "такими, как протянуты.";

    public override string Title => "Распред · пересчёт";

    public override string Subtitle => "Стенки протянуты → «остатки», «Распред», «Загрузочник»";

    /// <summary>Сетка распределения: столбцы разной высоты под общей чертой.</summary>
    public override string IconData =>
        "M3,16.6 L17,16.6 " +
        "M5.4,16.6 L5.4,11.2 " +
        "M9.1,16.6 L9.1,6.4 " +
        "M12.8,16.6 L12.8,8.8 " +
        "M16.5,16.6 L16.5,4.2";

    private const string AllOption = "Весь файл";
    private const string LinkOption = "Только link";
    private const string StockOption = "Только остатки";

    public override IReadOnlyList<string> RunOptions { get; } = new[] { AllOption, LinkOption, StockOption };

    public override string RunOptionsTitle => "ЧТО ПЕРЕСЧИТАТЬ";

    public override string RunOptionHint => Scope switch
    {
        RecalculateScope.Link =>
            "Обновили лист «link» - программа заново сверит с ним «Цены»: «Линк», «Проверка запретов», " +
            "даты и грейды, подсветка. Наценки, цены, «остатки», «Распред» и «Загрузочник» не трогаются.",
        RecalculateScope.Stock =>
            "Обновили «Остатки Н» - программа заново соберёт лист «остатки», переставит на него ссылки " +
            "«Распреда» и проверит «мин запас на Хаб» по свежим остаткам. Блоки, фотографии и «Цены» не трогаются.",
        _ =>
            "Наценки, аналоги, согласованные цены, проверки по стенкам и «link», листы «остатки», " +
            "«Распред» и «Загрузочник».",
    };

    /// <summary>Что пересчитывать. Обработчик читает это в момент запуска.</summary>
    public RecalculateScope Scope => SelectedRunOption switch
    {
        LinkOption => RecalculateScope.Link,
        StockOption => RecalculateScope.Stock,
        _ => RecalculateScope.All,
    };

    public override string ActionCaption => Scope switch
    {
        RecalculateScope.Link => "Пересчитать link",
        RecalculateScope.Stock => "Пересчитать остатки",
        _ => "Пересчитать распред",
    };

    public override string InitialHint =>
        "Выберите файл после подготовки - тот, в котором вы уже протянули колонки со стенками. " +
        "Понадобятся свежие листы «link», «Остатки Н», «КС», «Аналоги» и «Регламент наценок». " +
        "Листы книги программа не удаляет.";

    public override string SuccessMessage => Scope switch
    {
        RecalculateScope.Link => "Сверка с «link» пересчитана",
        RecalculateScope.Stock => "Остатки пересчитаны",
        _ => "Распред пересчитан",
    };

    public override string FailureMessage => "Не удалось пересчитать распред.";
}
