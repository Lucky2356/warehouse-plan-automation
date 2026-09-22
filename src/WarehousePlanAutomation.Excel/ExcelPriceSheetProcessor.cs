using System.Globalization;
using System.Runtime.InteropServices;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>Какой из двух этапов распреда выполняется.</summary>
public enum PriceStage
{
    /// <summary>
    /// Подготовка: «для цен» и «Цены» из инвойса и справочников. После неё аналитик
    /// протягивает колонки со стенками.
    /// </summary>
    Prepare,

    /// <summary>
    /// Пересчёт: наценки, аналоги, согласованные цены, проверки по стенкам и «link»,
    /// листы «остатки», «Распред» и «Загрузочник». Листы книги не удаляются.
    /// Работает по тому, что уже стоит на листе «Цены», и инвойс больше не читает.
    /// </summary>
    Recalculate,
}

/// <summary>Что пересчитывает второй этап распреда.</summary>
public enum RecalculateScope
{
    /// <summary>Весь файл: наценки, аналоги, цены, проверки, «остатки», «Распред», «Загрузочник».</summary>
    All,

    /// <summary>
    /// Только сверка с листом «link»: «Линк», «Проверка запретов» и подсветка на «Цены».
    /// Нужна, когда обновили лист «link», а остальное уже посчитано.
    /// </summary>
    Link,

    /// <summary>
    /// Только остатки: лист «остатки» заново из «Остатков Н», ссылки «Распреда» на него
    /// и «мин запас на Хаб» по свежим остаткам. Блоки и загрузочник не трогаются.
    /// </summary>
    Stock,
}

/// <summary>
/// Книга распреда. Работа разделена на два этапа, потому что между ними человек делает
/// то, чего программа сделать не может: протягивает колонки со стенками из файлов
/// в сетевых папках.
///
/// Исходный файл никогда не изменяется: все действия выполняются над копией.
/// Стенки программа только читает, сверяет и подсвечивает расхождения. Из сетевой папки
/// она читает одно - файлы согласования цен, и только если человек сам указал эту папку.
/// </summary>
public sealed class ExcelPriceSheetProcessor : IWorkbookProcessor
{
    private const int HeaderScanRows = 15;
    private const int ChunkRows = 50000;

    /// <summary>
    /// Насколько широкой может быть фотография на листе «Цены». Размер она сохраняет свой,
    /// это ограничение на случай, если в инвойс вставили картинку во весь экран.
    /// </summary>
    private const double PhotoWidthLimit = 200d;

    /// <summary>Высота строки текста в ячейке «артикул», которую фотография не закрывает.</summary>
    private const double ArticleTextHeight = 15d;

    /// <summary>
    /// Меньше этой высоты фотографии в строке «Цены» не на что смотреть. Столько же стоит
    /// у строк в примере распреда.
    /// </summary>
    private const double MinimumPhotoRowHeight = 63.75d;

    /// <summary>Сколько строк под итоговой просматривается в поиске формул, которые нужно растянуть.</summary>
    private const int BelowTableRows = 50;

    private readonly IAppLogger _logger;
    private readonly IDecisionPrompt? _prompt;
    private readonly PriceStage _stage;
    private readonly Func<DateTime> _nowProvider;
    private readonly Func<string?> _approvedPricesFolder;
    private readonly Func<RecalculateScope> _scope;

    /// <param name="approvedPricesFolder">
    /// Папка с файлами согласования цен. Спрашивается в момент запуска: человек мог выбрать
    /// её уже после того, как окно открылось. Пусто - «Согласованные цены» протягивает человек.
    /// </param>
    /// <param name="scope">Что пересчитывать на втором этапе. Спрашивается в момент запуска.</param>
    public ExcelPriceSheetProcessor(
        IAppLogger logger,
        IDecisionPrompt? prompt = null,
        PriceStage stage = PriceStage.Prepare,
        Func<DateTime>? nowProvider = null,
        Func<string?>? approvedPricesFolder = null,
        Func<RecalculateScope>? scope = null)
    {
        _logger = logger;
        _prompt = prompt;
        _stage = stage;
        _nowProvider = nowProvider ?? (() => DateTime.Now);
        _approvedPricesFolder = approvedPricesFolder ?? (() => null);
        _scope = scope ?? (() => RecalculateScope.All);
    }

    public Task<ProcessingResult> ProcessAsync(
        string sourceFilePath,
        IProgress<ProcessingStage>? progress,
        CancellationToken cancellationToken) =>
        StaTaskRunner.RunAsync(() => Process(sourceFilePath, progress, cancellationToken), cancellationToken);

    private ProcessingResult Process(
        string sourceFilePath,
        IProgress<ProcessingStage>? progress,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
        {
            throw new WarehousePlanException("Файл не найден: " + sourceFilePath);
        }

        if (!WorkbookFile.IsSupported(sourceFilePath))
        {
            throw new WarehousePlanException(
                "Неподдерживаемый формат файла: " + Path.GetExtension(sourceFilePath) +
                ". Ожидается книга Excel (" + string.Join(", ", WorkbookFile.SupportedExtensions) + ").");
        }

        var outputPath = OutputPath.Build(
            sourceFilePath,
            _nowProvider(),
            _stage == PriceStage.Prepare ? OutputFileName.PrepareMark : OutputFileName.DistributionMark);
        _logger.Information(
            (_stage == PriceStage.Prepare ? "Начало подготовки распреда." : "Начало пересчёта распреда.") +
            " Исходный файл: " + sourceFilePath);
        Report(progress, "Создание копии файла", 5);

        File.Copy(sourceFilePath, outputPath);
        _logger.Information("Создана копия: " + outputPath);

        try
        {
            var result = ProcessCopy(outputPath, sourceFilePath, progress, cancellationToken);
            _logger.Information(
                "Этап завершён. " +
                string.Join(", ", result.Counters.Select(c => c.Caption + ": " + c.Value)));
            return result;
        }
        catch (Exception ex)
        {
            _logger.Error("Обработка прервана, частичный результат удалён.", ex);
            OutputPath.TryDelete(outputPath, _logger);
            throw;
        }
    }

    private ProcessingResult ProcessCopy(
        string path,
        string sourcePath,
        IProgress<ProcessingStage>? progress,
        CancellationToken cancellationToken)
    {
        using var host = ExcelApplicationHost.Start(_logger);
        using var scope = new ComScope();

        object applicationObject = host.Application;
        dynamic application = applicationObject;
        dynamic? workbook = null;
        var closed = false;

        try
        {
            dynamic workbooks = scope.Track(application.Workbooks);
            workbook = scope.Track(workbooks.Open(path, 0));
            application.Calculation = ExcelConstants.XlCalculationManual;

            var sheets = ResolveSheets((object)workbook, scope, _stage);
            cancellationToken.ThrowIfCancellationRequested();

            var outcome = _stage == PriceStage.Prepare
                ? Prepare(applicationObject, sheets, path, sourcePath, progress, cancellationToken)
                : _scope() switch
                {
                    RecalculateScope.Link => RecalculateLink(applicationObject, sheets, path, progress),
                    RecalculateScope.Stock => RecalculateStock(applicationObject, sheets, path, progress),
                    _ => Recalculate(applicationObject, sheets, path, progress, cancellationToken),
                };

            Report(progress, "Сохранение файла", 96);
            workbook.Save();
            workbook.Close(true);
            closed = true;

            Report(progress, "Готово", 100);
            return outcome;
        }
        finally
        {
            if (workbook is not null && !closed)
            {
                try
                {
                    workbook.Close(false);
                }
                catch (COMException ex)
                {
                    _logger.Warning("Не удалось корректно закрыть книгу после ошибки.", ex);
                }
            }
        }
    }

    // ================= Этап 1: подготовка =================

    /// <summary>
    /// Из инвойса и справочников собираются листы «для цен» и «Цены». Проверки, которым
    /// нужны стенки, здесь не делаются: их колонки аналитик протягивает после этого этапа.
    /// </summary>
    private ProcessingResult Prepare(
        object applicationObject,
        PriceSheets sheets,
        string path,
        string sourcePath,
        IProgress<ProcessingStage>? progress,
        CancellationToken cancellationToken)
    {
        dynamic application = applicationObject;

        {
            Report(progress, "Чтение инвойса", 12);
            var listSeparator = ExcelSheetOperations.GetListSeparator(applicationObject);
            ExcelSheetOperations.ShowAllRows(sheets.Prices);

            var fileName = Path.GetFileNameWithoutExtension(sourcePath);
            var invoices = sheets.Invoices.Select(sheet => ReadInvoice(sheet, fileName)).ToList();
            var lines = invoices.Sum(invoice => invoice.Lines.Count);
            _logger.Information("Прочитано строк инвойса: " + lines + " на " + invoices.Count + " листе(ах).");

            if (lines == 0)
            {
                throw new WarehousePlanException(
                    "На листе «" + PriceSchema.InvoiceSheet + "» не найдено ни одной строки товара. " +
                    "Проверьте, что в таблице есть колонки «штрих-код», «количество», «цена» и «сумма».");
            }

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Чтение справочников", 28);

            var barcodes = invoices
                .SelectMany(invoice => invoice.Lines)
                .Select(line => line.Barcode)
                .ToHashSet(StringComparer.Ordinal);

            MoveSummaryBarcodeFirst(sheets.SummaryPrice);
            var priceList = ReadSummaryPrice(sheets.SummaryPrice, barcodes);
            var codes = priceList.Values
                .Select(row => row.Code)
                .Where(code => code.Length > 0)
                .ToHashSet(StringComparer.Ordinal);
            var subgroups = ReadSubgroups(sheets.DivisionPrice, codes);

            _logger.Information(
                "Найдено в прайсе: " + priceList.Count + " из " + barcodes.Count +
                ", подгрупп " + subgroups.Count + ".");

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Сборка строк листа «Цены»", 45);
            var plan = PriceRowBuilder.Build(invoices, priceList, subgroups);
            plan = KeepOneSector(plan, sourcePath);

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Заполнение листа «для цен»", 55);
            FillForPrices(applicationObject, sheets.ForPrices, plan.Supplies);

            Report(progress, "Заполнение листа «Цены»", 65);
            var layout = ReadPricesLayout(sheets.Prices);
            layout = Resize(applicationObject, sheets.Prices, layout, plan.Rows.Count, listSeparator);
            NormalizeDataRows(applicationObject, sheets.Prices, layout);
            WriteRows(sheets.Prices, layout, plan.Rows);
            FixSupplyReferences(sheets.Prices, layout, plan.Rows);
            WriteTotals(sheets.Prices, layout, plan.Rows);

            Report(progress, "Перенос фотографий из инвойса", 72);
            var photos = CopyInvoicePhotos(
                applicationObject, sheets.Invoices, invoices, sheets.Prices, layout, plan.Rows);

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Пересчёт формул", 78);
            application.Calculation = ExcelConstants.XlCalculationAutomatic;
            application.CalculateFull();

            Report(progress, "Разметка того, что нужно проверить", 88);
            IReadOnlyList<PriceRowCheck> checks = plan.Rows.Select(PriceChecks.Prepare).ToList();
            ApplyPreparedChecks(sheets.Prices, layout, checks);
            MarkDuplicateBarcodes(sheets.Prices, layout, plan.Rows);
            checks = Locate(sheets.Prices, layout, checks);

            return BuildPrepareResult(path, plan, checks, photos);
        }
    }

    private static ProcessingResult BuildPrepareResult(
        string path,
        PriceSheetPlan plan,
        IReadOnlyList<PriceRowCheck> checks,
        PhotoOutcome photos)
    {
        var notFound = plan.Rows.Count(row => row.Reference is null);
        var highlighted = checks.Sum(check => check.Highlight.Count);
        var missingPhotos = plan.Rows.Count - photos.Copied;

        var warnings = plan.Warnings
            .Concat(checks.SelectMany(check => check.Warnings))
            .Concat(photos.Warnings)
            .ToList();

        return new ProcessingResult(
            path,
            new[]
            {
                new ProcessingCounter("строк заполнено", plan.Rows.Count),
                new ProcessingCounter("поставок", plan.Supplies.Count),
                new ProcessingCounter("нет в прайсе", notFound, notFound > 0),
                new ProcessingCounter("фото перенесено", photos.Copied, missingPhotos > 0),
                new ProcessingCounter("подсвечено ячеек", highlighted, highlighted > 0),
            },
            warnings);
    }

    /// <summary>
    /// Один файл - один сектор. Сектор виден только после подтягивания из «Сводного прайса»,
    /// поэтому лишние строки отсекаются здесь. Какой сектор нужен, программа узнаёт
    /// из названия файла; если не узнала - спрашивает.
    /// </summary>
    private PriceSheetPlan KeepOneSector(PriceSheetPlan plan, string sourcePath)
    {
        var sectors = SectorSelection.Sectors(plan.Rows);
        if (sectors.Count < 2)
        {
            return plan;
        }

        var matched = SectorSelection.MatchFileName(sectors, sourcePath);
        if (matched.Count == 1)
        {
            _logger.Information(
                "В инвойсе секторы: " + string.Join(", ", sectors) +
                ". По названию файла выбран «" + matched[0] + "».");
            return PriceRowBuilder.KeepSector(plan, matched[0]);
        }

        var answer = _prompt?.Choose(new DecisionRequest(
            "В инвойсе несколько секторов. Какой считать в этом файле?",
            "Один распред - один сектор. Остальные строки в файл не попадут, для них нужен " +
            "отдельный запуск на той же исходной книге. По названию файла сектор " +
            (matched.Count == 0 ? "не узнался." : "узнаётся неоднозначно."),
            sectors.ToList()));

        if (answer is null)
        {
            var warnings = plan.Warnings.ToList();
            warnings.Add(new ProcessingWarning(
                "В инвойсе несколько секторов (" + string.Join(", ", sectors) +
                "), выбор не сделан - в файл попали все. Распред считается по одному сектору.",
                PriceSchema.PricesSheet));
            return new PriceSheetPlan(plan.Rows, plan.Supplies, warnings);
        }

        _logger.Information("Выбран сектор «" + answer + "».");
        return PriceRowBuilder.KeepSector(plan, answer);
    }

    // ================= Этап 2: пересчёт =================

    /// <summary>
    /// Работает по тому, что уже стоит на листе «Цены»: колонки со стенками протянуты
    /// человеком. Инвойс и справочники на этом этапе не читаются.
    /// </summary>
    private ProcessingResult Recalculate(
        object applicationObject,
        PriceSheets sheets,
        string path,
        IProgress<ProcessingStage>? progress,
        CancellationToken cancellationToken)
    {
        dynamic application = applicationObject;

        Report(progress, "Чтение листа «Цены»", 12);
        ExcelSheetOperations.ShowAllRows(sheets.Prices);

        var layout = ReadPricesLayout(sheets.Prices);
        var rows = ReadPriceRows(sheets.Prices, layout);
        _logger.Information("На листе «Цены» строк: " + rows.Count + ".");

        if (rows.Count == 0)
        {
            throw new WarehousePlanException(
                "На листе «" + PriceSchema.PricesSheet + "» нет ни одной строки со штрихкодом. " +
                "Сначала выполните подготовку распреда.");
        }

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Наценки по регламенту", 20);
        var markupRules = MarkupSheetReader.Read(
            ExcelSheetOperations.ReadGrid(sheets.Markup, withFormulas: false));
        var markups = MarkupAssignment.Apply(rows, new MarkupResolver(markupRules, _prompt));
        rows = markups.Rows;
        WriteMarkups(sheets.Prices, layout, rows);

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Аналоги по листу «Аналоги»", 21);
        var extraWarnings = new List<ProcessingWarning>();
        var analogues = FillGoogleAnalogues(sheets.Analogues, sheets.Prices, layout, extraWarnings);

        ApprovedPricesOutcome? approved = null;
        var folder = _approvedPricesFolder()?.Trim();
        if (!string.IsNullOrEmpty(folder))
        {
            Report(progress, "Согласованные цены из файлов согласования", 22);
            if (layout.Headers.TryGet(PriceSchema.Prices.ApprovedPrice, out var approvedColumn))
            {
                approved = new ExcelApprovedPricesStage(_logger).Run(
                    applicationObject, folder, sheets.Prices, layout.FirstDataRow, approvedColumn, rows, cancellationToken);
            }
            else
            {
                extraWarnings.Add(new ProcessingWarning(
                    "На листе «" + PriceSchema.PricesSheet + "» нет колонки «" + PriceSchema.Prices.ApprovedPrice +
                    "» - согласованные цены из папки некуда записать.",
                    PriceSchema.PricesSheet));
            }
        }
        else
        {
            _logger.Information("Файл или папка согласования цен не выбраны: «Согласованные цены» берутся такими, как протянуты.");
        }

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Пересчёт формул", 24);
        application.Calculation = ExcelConstants.XlCalculationAutomatic;
        application.CalculateFull();

        Report(progress, "Проверка стенок и сверка с листом «link»", 40);
        var states = ReadRowStates(sheets.Prices, layout, rows);
        var links = ReadLink(
            sheets.Link,
            states.Select(s => TextUtils.NormalizeKey(s.Acr)).ToHashSet(StringComparer.Ordinal));
        var checks = RunChecks(rows, states, links);
        ApplyChecks(sheets.Prices, layout, checks);
        MarkDuplicateBarcodes(sheets.Prices, layout, rows);
        checks = Locate(sheets.Prices, layout, checks);

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Заполнение листов «остатки» и «Распред»", 60);
        var stage = new ExcelDistributionStage(_logger, _nowProvider);
        var distribution = stage.Run(
            applicationObject,
            sheets.StockSource,
            sheets.Stock,
            sheets.Distribution,
            sheets.Seasonality,
            sheets.Prices,
            layout.FirstDataRow,
            layout.Columns,
            rows,
            cancellationToken);

        Report(progress, "Заполнение листа «Загрузочник»", 80);
        var loaderStage = new ExcelLoaderStage(_logger);
        var loader = loaderStage.Run(applicationObject, sheets.Loader, rows);

        Report(progress, "Пересчёт распределения", 88);
        application.CalculateFull();
        distribution = stage.Finish(sheets.Distribution, distribution, () => application.CalculateFull());

        cancellationToken.ThrowIfCancellationRequested();
        application.CalculateFull();

        // Проверка загрузочника читается последней: она сходится только тогда, когда
        // посчитан и распред, и обнулённый «мин запас на Хаб».
        loader = loaderStage.Verify(sheets.Loader, loader);

        return BuildRecalculateResult(
            path, rows, checks, markups, distribution, loader, analogues, approved, extraWarnings);
    }

    /// <summary>
    /// Пересчёт только сверки с листом «link»: «Линк», «Проверка запретов», подсветка
    /// и повторы штрихкодов на листе «Цены». Наценки, цены и остальные листы не трогаются.
    /// </summary>
    private ProcessingResult RecalculateLink(
        object applicationObject,
        PriceSheets sheets,
        string path,
        IProgress<ProcessingStage>? progress)
    {
        dynamic application = applicationObject;
        _logger.Information("Пересчитывается только сверка с листом «link».");

        Report(progress, "Чтение листа «Цены»", 12);
        ExcelSheetOperations.ShowAllRows(sheets.Prices);
        var layout = ReadPricesLayout(sheets.Prices);
        var rows = RequireRows(ReadPriceRows(sheets.Prices, layout));

        Report(progress, "Пересчёт формул", 30);
        application.Calculation = ExcelConstants.XlCalculationAutomatic;
        application.CalculateFull();

        Report(progress, "Сверка с листом «link»", 60);
        var states = ReadRowStates(sheets.Prices, layout, rows);
        var links = ReadLink(
            sheets.Link,
            states.Select(s => TextUtils.NormalizeKey(s.Acr)).ToHashSet(StringComparer.Ordinal));
        var checks = RunChecks(rows, states, links);
        ApplyChecks(sheets.Prices, layout, checks);
        MarkDuplicateBarcodes(sheets.Prices, layout, rows);
        checks = Locate(sheets.Prices, layout, checks);

        Report(progress, "Пересчёт формул", 88);
        application.CalculateFull();

        var missingLink = checks.Count(check => check.LinkValue == PriceChecks.LinkMissing);
        var highlighted = checks.Sum(check => check.Highlight.Count);

        return new ProcessingResult(
            path,
            new List<ProcessingCounter>
            {
                new("строк «Цены»", rows.Count),
                new("нет в Линке", missingLink, missingLink > 0),
                new("подсвечено ячеек", highlighted, highlighted > 0),
            },
            checks.SelectMany(check => check.Warnings).ToList());
    }

    /// <summary>
    /// Пересчёт только остатков: лист «остатки» из свежих «Остатков Н», ссылки «Распреда»
    /// на него, «мин запас на Хаб» по правилу 30 %, сортировка РТТ и проверка загрузочника.
    /// Наценки, проверки, блоки, фотографии и сезонность остаются как есть.
    /// </summary>
    private ProcessingResult RecalculateStock(
        object applicationObject,
        PriceSheets sheets,
        string path,
        IProgress<ProcessingStage>? progress)
    {
        dynamic application = applicationObject;
        _logger.Information("Пересчитываются только остатки.");

        Report(progress, "Чтение листа «Цены»", 12);
        ExcelSheetOperations.ShowAllRows(sheets.Prices);
        var layout = ReadPricesLayout(sheets.Prices);
        var rows = RequireRows(ReadPriceRows(sheets.Prices, layout));

        Report(progress, "Заполнение листа «остатки»", 40);
        var stage = new ExcelDistributionStage(_logger, _nowProvider);
        var distribution = stage.RunStockOnly(sheets.StockSource, sheets.Stock, sheets.Distribution, rows);

        Report(progress, "Пересчёт распределения", 70);
        application.Calculation = ExcelConstants.XlCalculationAutomatic;
        application.CalculateFull();
        distribution = stage.Finish(sheets.Distribution, distribution, () => application.CalculateFull());
        application.CalculateFull();

        var loader = new ExcelLoaderStage(_logger).Verify(
            sheets.Loader, new LoaderOutcome(0, null, Array.Empty<ProcessingWarning>()));

        return new ProcessingResult(
            path,
            new List<ProcessingCounter>
            {
                new("АЦР в остатках", distribution.StockColumns),
                new("блоков «Распред»", distribution.Blocks),
                new("обнулено «мин запас на Хаб»", distribution.ZeroedHubMinimums),
            },
            distribution.Warnings.Concat(loader.Warnings).ToList());
    }

    private static IReadOnlyList<PriceRowValues> RequireRows(IReadOnlyList<PriceRowValues> rows)
    {
        if (rows.Count == 0)
        {
            throw new WarehousePlanException(
                "На листе «" + PriceSchema.PricesSheet + "» нет ни одной строки со штрихкодом. " +
                "Сначала выполните подготовку распреда.");
        }

        return rows;
    }

    private static ProcessingResult BuildRecalculateResult(
        string path,
        IReadOnlyList<PriceRowValues> rows,
        IReadOnlyList<PriceRowCheck> checks,
        MarkupPlan markups,
        DistributionOutcome distribution,
        LoaderOutcome loader,
        int analogues,
        ApprovedPricesOutcome? approved,
        IReadOnlyList<ProcessingWarning> extraWarnings)
    {
        var missingLink = checks.Count(check => check.LinkValue == PriceChecks.LinkMissing);
        var highlighted = checks.Sum(check => check.Highlight.Count);
        var noMarkup = rows.Count(row => row.Markup is null);

        var warnings = checks
            .SelectMany(check => check.Warnings)
            .Concat(markups.Warnings)
            .Concat(extraWarnings)
            .Concat(approved?.Warnings ?? Array.Empty<ProcessingWarning>())
            .Concat(distribution.Warnings)
            .Concat(loader.Warnings)
            .ToList();

        var counters = new List<ProcessingCounter>
        {
            new("строк «Цены»", rows.Count),
            new("без наценки", noMarkup, noMarkup > 0),
            new("нет в Линке", missingLink, missingLink > 0),
            new("с аналогом", analogues),
        };

        if (approved is not null)
        {
            counters.Add(new ProcessingCounter("согласованных цен", approved.Filled, approved.Missing > 0));
        }

        counters.AddRange(new[]
        {
            new ProcessingCounter("подсвечено ячеек", highlighted, highlighted > 0),
            new ProcessingCounter("АЦР в остатках", distribution.StockColumns),
            new ProcessingCounter("блоков «Распред»", distribution.Blocks),
            new ProcessingCounter("колонок загрузочника", loader.Columns, loader.Check != true),
        });

        return new ProcessingResult(path, counters, warnings);
    }

    /// <summary>
    /// «Аналог гугл» по листу «Аналоги»: название аналога либо «нет». Пишется значениями -
    /// вручную колонку тоже в конце вставляют как значения.
    /// Возвращает, у скольких строк аналог нашёлся.
    /// </summary>
    private int FillGoogleAnalogues(
        object? analoguesSheet,
        object pricesSheet,
        PricesLayout layout,
        List<ProcessingWarning> warnings)
    {
        if (analoguesSheet is null)
        {
            warnings.Add(new ProcessingWarning(
                "В книге нет листа «" + PriceSchema.AnaloguesSheet + "» - «" + PriceSchema.Prices.GoogleAnalogue +
                "» не заполнен программой и сверяется таким, как стоит в книге.",
                PriceSchema.PricesSheet));
            return 0;
        }

        var sheetName = ExcelSheetOperations.GetSheetName(analoguesSheet);
        var grid = ExcelSheetOperations.ReadGrid(analoguesSheet, withFormulas: false);

        HeaderMap headers;
        try
        {
            headers = HeaderResolver.Resolve(grid, sheetName, PriceSchema.Analogues.Specs);
        }
        catch (WorkbookValidationException ex)
        {
            warnings.Add(new ProcessingWarning(
                string.Join("; ", ex.Problems) + " - «" + PriceSchema.Prices.GoogleAnalogue + "» не заполнен программой.",
                sheetName,
                sheetName));
            return 0;
        }

        var pairs = new List<AnaloguePair>();
        for (var row = headers.HeaderRow + 1; row <= grid.LastRow; row++)
        {
            pairs.Add(new AnaloguePair(
                grid.Text(row, headers[PriceSchema.Analogues.Vsp]),
                grid.Text(row, headers[PriceSchema.Analogues.Analogue])));
        }

        var vspColumn = layout.Headers[PriceSchema.Prices.Vsp];
        var vspGrid = ExcelSheetOperations.ReadBlock(
            pricesSheet, layout.FirstDataRow, layout.LastDataRow, vspColumn, vspColumn, withFormulas: false);
        var vsp = Enumerable.Range(layout.FirstDataRow, layout.DataRowCount)
            .Select(row => vspGrid.Text(row, vspColumn))
            .ToList();

        var lookup = AnalogueLookup.Build(pairs, vsp);
        var values = vsp.Select(value => (object?)AnalogueLookup.ValueFor(lookup, value)).ToList();
        ExcelSheetOperations.SetColumnValues(
            pricesSheet, layout.FirstDataRow, layout.Headers[PriceSchema.Prices.GoogleAnalogue], values);

        var withAnalogue = values.Count(value => !Equals(value, AnalogueLookup.None));
        _logger.Information(
            "«Аналог гугл»: строк на листе «" + sheetName + "» " + pairs.Count + ", аналог нашёлся у " +
            withAnalogue + " из " + values.Count + ".");
        return withAnalogue;
    }

    // ================= Поиск листов =================

    private sealed record PriceSheets(
        IReadOnlyList<object> Invoices,
        object ForPrices,
        object Prices,
        object SummaryPrice,
        object DivisionPrice,
        object Markup,
        object Link,
        object StockSource,
        object Stock,
        object Distribution,
        object Loader,
        object Seasonality,
        object? Analogues);

    /// <summary>
    /// Ищет только те листы, которые нужны этому этапу: на пересчёте не должно быть
    /// требования к инвойсу, а на подготовке - к «Остаткам Н».
    /// </summary>
    private static PriceSheets ResolveSheets(object workbook, ComScope scope, PriceStage stage)
    {
        var problems = new List<string>();
        var prepare = stage == PriceStage.Prepare;

        // Инвойс может прийти двумя листами - оба правильные, поэтому берутся все.
        IReadOnlyList<object> invoices = prepare
            ? ExcelSheetOperations.FindSheets(workbook, PriceSchema.InvoiceSheet, scope)
            : Array.Empty<object>();
        if (prepare && invoices.Count == 0)
        {
            problems.Add("в книге нет листа, название которого начинается с «" +
                         PriceSchema.InvoiceSheet + "»");
        }

        object? Single(string name)
        {
            var lookup = ExcelSheetOperations.FindSheet(workbook, name, scope);
            if (lookup.Sheet is not null)
            {
                return lookup.Sheet;
            }

            problems.Add(lookup.Candidates.Count == 0
                ? "в книге нет листа, название которого начинается с «" + name + "»"
                : "в книге несколько листов с названием, начинающимся с «" + name + "»: " +
                  string.Join(", ", lookup.Candidates) + " - оставьте один");
            return null;
        }

        object? Needed(string name, bool required) => required ? Single(name) : null;

        var prices = Single(PriceSchema.PricesSheet);
        var forPrices = Needed(PriceSchema.ForPricesSheet, prepare);
        var summary = Needed(PriceSchema.SummaryPriceSheet, prepare);
        var division = Needed(PriceSchema.DivisionPriceSheet, prepare);
        // Наценки ставит пересчёт: на подготовке подгруппы ещё только собираются,
        // и спрашивать о наценке для каждой из них рано.
        var markup = Needed(PriceSchema.MarkupSheet, !prepare);
        var link = Needed(PriceSchema.LinkSheet, !prepare);
        var stockSource = Needed(PriceSchema.StockSourceSheet, !prepare);
        var stock = Needed(PriceSchema.StockSheet, !prepare);
        var distribution = Needed(PriceSchema.DistributionSheet, !prepare);
        var loader = Needed(PriceSchema.LoaderSheet, !prepare);
        var seasonality = Needed(PriceSchema.SeasonalitySheet, !prepare);

        // Без «Аналогов» пересчёт идёт: «Аналог гугл» тогда сверяется таким, как стоит в книге.
        var analogues = prepare
            ? null
            : ExcelSheetOperations.FindSheet(workbook, PriceSchema.AnaloguesSheet, scope).Sheet;

        if (problems.Count > 0)
        {
            throw new WorkbookValidationException(problems);
        }

        return new PriceSheets(
            invoices, forPrices!, prices!, summary!, division!, markup!, link!,
            stockSource!, stock!, distribution!, loader!, seasonality!, analogues);
    }

    // ================= Чтение =================

    private static InvoiceSheet ReadInvoice(object sheet, string fileName) =>
        InvoiceSheetReader.Read(
            ExcelSheetOperations.ReadGrid(sheet, withFormulas: false),
            ExcelSheetOperations.GetSheetName(sheet),
            fileName);

    /// <summary>
    /// Колонка «Штрихкод» на «Сводном прайсе» выносится в начало листа - как по инструкции
    /// перед заполнением «Цен»: по ней ищут ВПР, а ВПР ищет только в первой колонке диапазона.
    ///
    /// Перенос без буфера обмена: в начало вставляется пустая колонка, «Штрихкод» переезжает
    /// в неё вырезанием с указанным местом, опустевшая колонка удаляется. Ссылки на перенесённые
    /// ячейки при вырезании идут за ними, как при ручном «Вырезать - Вставить вырезанные ячейки».
    /// </summary>
    private void MoveSummaryBarcodeFirst(object sheet)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(sheet);
        var headerGrid = ExcelSheetOperations.ReadBlock(
            sheet,
            bounds.FirstRow,
            Math.Min(bounds.FirstRow + HeaderScanRows - 1, bounds.LastRow),
            1,
            bounds.LastColumn,
            withFormulas: false);

        var headers = PriceListReader.ResolveSummaryHeaders(headerGrid);
        var column = headers[PriceSchema.SummaryPrice.Barcode];
        if (column == 1)
        {
            return;
        }

        ExcelSheetOperations.InsertColumns(sheet, 1, 1);
        ExcelSheetOperations.MoveColumn(sheet, column + 1, 1);
        ExcelSheetOperations.DeleteColumns(sheet, column + 1, 1);

        var moved = TextUtils.Normalize(ExcelSheetOperations.GetValue(sheet, headers.HeaderRow, 1)?.ToString());
        _logger.Information(
            "На листе «" + PriceSchema.SummaryPriceSheet + "» колонка «" + moved + "» вынесена в начало из колонки " +
            ExcelColumn.ToLetters(column) + ".");
    }

    /// <summary>
    /// «Сводный прайс» - сотни тысяч строк, из которых нужны единицы. Лист читается
    /// порциями и только по нужным колонкам, в память попадают лишь совпавшие строки.
    /// </summary>
    private Dictionary<string, PriceListRow> ReadSummaryPrice(object sheet, IReadOnlySet<string> barcodes)
    {
        var result = new Dictionary<string, PriceListRow>(StringComparer.Ordinal);
        ReadByChunks(
            sheet,
            PriceListReader.ResolveSummaryHeaders,
            PriceSchema.SummaryPrice.Barcode,
            (chunk, headers) => PriceListReader.ReadSummary(chunk, headers, barcodes, result),
            () => result.Count >= barcodes.Count);
        return result;
    }

    private Dictionary<string, string> ReadSubgroups(object sheet, IReadOnlySet<string> codes)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        ReadByChunks(
            sheet,
            PriceListReader.ResolveDivisionHeaders,
            PriceSchema.DivisionPrice.Code,
            (chunk, headers) => PriceListReader.ReadSubgroups(chunk, headers, codes, result),
            () => result.Count >= codes.Count);
        return result;
    }

    private Dictionary<string, LinkEntry> ReadLink(object sheet, IReadOnlySet<string> acr)
    {
        var rows = new Dictionary<string, List<LinkRow>>(StringComparer.Ordinal);
        ReadByChunks(
            sheet,
            LinkSheetReader.ResolveHeaders,
            PriceSchema.Link.Acr,
            (chunk, headers) => LinkSheetReader.Read(chunk, headers, acr, rows),
            () => false);

        return rows.ToDictionary(
            pair => pair.Key,
            pair => LinkSheetReader.Summarize(pair.Key, pair.Value),
            StringComparer.Ordinal);
    }

    /// <summary>
    /// Общий обход большого листа: строка заголовков ищется по первым строкам, дальше лист
    /// читается порциями и только по нужным колонкам. <paramref name="isComplete"/> позволяет
    /// остановиться, как только всё искомое найдено, и не дочитывать сотню тысяч строк.
    /// </summary>
    private void ReadByChunks(
        object sheet,
        Func<SheetGrid, HeaderMap> resolveHeaders,
        string keyColumn,
        Action<SheetGrid, HeaderMap> readChunk,
        Func<bool> isComplete)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(sheet);
        var headerGrid = ExcelSheetOperations.ReadBlock(
            sheet,
            bounds.FirstRow,
            Math.Min(bounds.FirstRow + HeaderScanRows - 1, bounds.LastRow),
            bounds.FirstColumn,
            bounds.LastColumn,
            withFormulas: false);

        var headers = resolveHeaders(headerGrid);
        var firstColumn = headers.Columns.Values.Min();
        var lastColumn = headers.Columns.Values.Max();

        var lastRow = Math.Min(
            bounds.LastRow,
            ExcelSheetOperations.GetLastFilledRow(sheet, headers[keyColumn], bounds.LastRow));

        for (var start = headers.HeaderRow + 1; start <= lastRow; start += ChunkRows)
        {
            if (isComplete())
            {
                return;
            }

            var end = Math.Min(start + ChunkRows - 1, lastRow);
            var chunk = ExcelSheetOperations.ReadBlock(
                sheet, start, end, firstColumn, lastColumn, withFormulas: false);
            readChunk(chunk, headers);
        }
    }

    // ================= Лист «для цен» =================

    /// <summary>
    /// Заполняет «Поставку», «Количество» и «Сумму закупа». «Рублевая оплата», «Курс»
    /// и «Логистика» не трогаются: их вносит человек, и от них зависят формулы
    /// «% логистики» и «Фактический курс».
    ///
    /// Колонок на листе становится ровно столько, сколько поставок. Недостающая создаётся
    /// копией первой - вместе с формулами «Курс» и «% логистики»; лишняя очищается целиком.
    /// В прошлом файле поставок могло быть две, а в этом одна: её цифры остались бы на листе
    /// и попали бы в отчёт.
    /// </summary>
    private void FillForPrices(object application, object sheet, IReadOnlyList<SupplySummary> supplies)
    {
        var grid = ExcelSheetOperations.ReadGrid(sheet, withFormulas: true);

        int? FindRow(string label)
        {
            for (var row = grid.FirstRow; row <= grid.LastRow; row++)
            {
                if (TextUtils.StartsWithKey(grid.Text(row, grid.FirstColumn), label))
                {
                    return row;
                }
            }

            return null;
        }

        var supplyRow = FindRow(PriceSchema.ForPrices.Supply);
        var quantityRow = FindRow(PriceSchema.ForPrices.Quantity);
        var amountRow = FindRow(PriceSchema.ForPrices.Amount);

        if (supplyRow is null || quantityRow is null || amountRow is null)
        {
            throw new WorkbookValidationException(new[]
            {
                "на листе «" + PriceSchema.ForPricesSheet + "» не найдены строки «Поставка», " +
                "«Количество» и «Сумма закупа»",
            });
        }

        var labelColumn = grid.FirstColumn;
        var firstValue = labelColumn + 1;
        var firstRow = Math.Min(supplyRow.Value, Math.Min(quantityRow.Value, amountRow.Value));
        var lastRow = LastLabelRow(grid, labelColumn, firstRow);

        ResizeForPrices(application, sheet, grid, firstRow, lastRow, firstValue, supplies.Count);

        foreach (var supply in supplies)
        {
            var column = firstValue + supply.InvoiceIndex;
            ExcelSheetOperations.SetValue(sheet, supplyRow.Value, column, supply.Supply);
            ExcelSheetOperations.SetValue(sheet, quantityRow.Value, column, supply.Quantity);
            ExcelSheetOperations.SetValue(sheet, amountRow.Value, column, supply.Amount);
            _logger.Debug(
                "«для цен», колонка " + ExcelColumn.ToLetters(column) + ": поставка " + supply.Supply +
                ", количество " + supply.Quantity + ", сумма " + supply.Amount + ".");
        }
    }

    /// <summary>Последняя строка блока: ниже неё в колонке подписей уже пусто.</summary>
    private static int LastLabelRow(SheetGrid grid, int labelColumn, int firstRow)
    {
        var last = firstRow;
        for (var row = firstRow; row <= grid.LastRow; row++)
        {
            if (TextUtils.Normalize(grid.Text(row, labelColumn)).Length > 0)
            {
                last = row;
            }
        }

        return last;
    }

    /// <summary>
    /// Доводит число колонок «для цен» до числа поставок. Новая колонка - копия первой:
    /// так в ней оказываются формулы «Курс» и «% логистики». Значения, которые вносит
    /// человек, из копии убираются: они относятся к другой поставке.
    /// </summary>
    private void ResizeForPrices(
        object application,
        object sheet,
        SheetGrid grid,
        int firstRow,
        int lastRow,
        int firstValue,
        int wanted)
    {
        var lastFilled = firstValue - 1;
        for (var column = firstValue; column <= grid.LastColumn; column++)
        {
            for (var row = firstRow; row <= lastRow; row++)
            {
                if (TextUtils.Normalize(grid.Text(row, column)).Length > 0)
                {
                    lastFilled = column;
                    break;
                }
            }
        }

        for (var column = firstValue + Math.Max(lastFilled - firstValue + 1, 0); column < firstValue + wanted; column++)
        {
            ExcelSheetOperations.CopyRange(
                application, sheet, firstRow, firstValue, lastRow, firstValue, firstRow, column);

            // В копии остались цифры первой поставки. Формулы остаются, всё остальное - нет.
            for (var row = firstRow; row <= lastRow; row++)
            {
                if (!grid.HasFormula(row, firstValue))
                {
                    ExcelSheetOperations.ClearValue(sheet, row, column);
                }
            }

            _logger.Information(
                "На листе «" + PriceSchema.ForPricesSheet + "» добавлена колонка " +
                ExcelColumn.ToLetters(column) + " - копия первой.");
        }

        for (var column = firstValue + wanted; column <= lastFilled; column++)
        {
            ExcelSheetOperations.ClearRange(sheet, firstRow, lastRow, column, column);
            _logger.Information(
                "На листе «" + PriceSchema.ForPricesSheet + "» очищена лишняя колонка " +
                ExcelColumn.ToLetters(column) + ": поставок " + wanted + ".");
        }
    }

    // ================= Лист «Цены» =================

    /// <summary>
    /// Разметка листа «Цены»: строка-образец, строки данных и итоговая строка под ними.
    /// </summary>
    private sealed record PricesLayout(
        HeaderMap Headers,
        int FirstDataRow,
        int DataRowCount,
        int TotalsRow,
        ColumnRange Columns)
    {
        public int LastDataRow => FirstDataRow + DataRowCount - 1;
    }

    private static PricesLayout ReadPricesLayout(object sheet)
    {
        var grid = ExcelSheetOperations.ReadGrid(sheet, withFormulas: true);
        var headers = WithOptionalColumns(
            grid, HeaderResolver.Resolve(grid, PriceSchema.PricesSheet, PriceSchema.Prices.Specs));

        var barcodeColumn = headers[PriceSchema.Prices.BarcodeCopy];
        var unitsColumn = headers[PriceSchema.Prices.Units];
        var amountColumn = headers[PriceSchema.Prices.PurchaseAmount];

        var firstData = headers.HeaderRow + 1;
        var totalsRow = 0;
        var lastFilled = 0;

        for (var row = firstData; row <= grid.LastRow; row++)
        {
            var hasBarcode = TextUtils.Normalize(grid.Text(row, barcodeColumn)).Length > 0;
            if (hasBarcode)
            {
                lastFilled = row;
                continue;
            }

            // Итоговая строка узнаётся так же, как её видит человек: штрихкода нет,
            // а количество и сумма стоят.
            var hasTotals = TextUtils.Normalize(grid.Text(row, unitsColumn)).Length > 0 ||
                            TextUtils.Normalize(grid.Text(row, amountColumn)).Length > 0;
            if (hasTotals)
            {
                totalsRow = row;
                break;
            }
        }

        if (totalsRow == 0)
        {
            totalsRow = Math.Max(lastFilled, firstData) + 1;
        }

        // Вставка и удаление строк идут по всей ширине занятого диапазона, а не только
        // по найденным колонкам: справа от последнего заголовка на листе стоят колонки,
        // которые в схеме не описаны («Коментарий ГН»), и сдвигаться они должны вместе
        // со своей строкой.
        return new PricesLayout(
            headers,
            firstData,
            Math.Max(totalsRow - firstData, 1),
            totalsRow,
            new ColumnRange(
                Math.Min(headers.Columns.Values.Min(), grid.FirstColumn),
                Math.Max(headers.Columns.Values.Max(), grid.LastColumn)));
    }

    /// <summary>
    /// Дописывает к найденным колонкам необязательные: «МП», «Согласованные цены». Ищутся они
    /// в той же строке заголовков полным совпадением и только среди колонок, не занятых схемой.
    /// </summary>
    private static HeaderMap WithOptionalColumns(SheetGrid grid, HeaderMap headers)
    {
        var columns = headers.Columns.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        var used = columns.Values.ToHashSet();

        foreach (var spec in PriceSchema.Prices.OptionalSpecs)
        {
            for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
            {
                if (!used.Contains(column) &&
                    spec.Aliases.Contains(TextUtils.NormalizeKey(grid.Text(headers.HeaderRow, column))))
                {
                    columns[spec.DisplayName] = column;
                    used.Add(column);
                    break;
                }
            }
        }

        return new HeaderMap(headers.HeaderRow, columns);
    }

    /// <summary>
    /// Доводит число строк данных до числа штрихкодов. Новые строки создаются копией
    /// первой строки данных - так они наследуют все её формулы, форматы и оформление,
    /// и ничего «протягивать» отдельно не нужно.
    ///
    /// Вставка и удаление идут внутри колонок таблицы: справа от неё на листе может
    /// стоять что угодно, и оно не должно уезжать вместе со строками.
    /// </summary>
    private PricesLayout Resize(
        object application,
        object sheet,
        PricesLayout layout,
        int wanted,
        string listSeparator)
    {
        if (wanted <= 0 || wanted == layout.DataRowCount)
        {
            return layout;
        }

        if (wanted > layout.DataRowCount)
        {
            var count = wanted - layout.DataRowCount;
            ExcelSheetOperations.InsertCopiedRows(
                application, sheet, layout.FirstDataRow, layout.TotalsRow, count, layout.Columns);
            _logger.Information("В лист «Цены» добавлено строк: " + count + ".");
        }
        else
        {
            var rows = Enumerable.Range(layout.FirstDataRow + wanted, layout.DataRowCount - wanted).ToList();
            ExcelSheetOperations.DeleteRows(sheet, rows, listSeparator, layout.Columns);
            _logger.Information("Из листа «Цены» удалено лишних строк: " + rows.Count + ".");
        }

        return layout with { DataRowCount = wanted, TotalsRow = layout.FirstDataRow + wanted };
    }

    private void WriteRows(object sheet, PricesLayout layout, IReadOnlyList<PriceRowValues> rows)
    {
        var headers = layout.Headers;

        void Write(string column, Func<PriceRowValues, object?> select) =>
            ExcelSheetOperations.SetColumnValues(
                sheet, layout.FirstDataRow, headers[column], rows.Select(select).ToList());

        Write(PriceSchema.Prices.PurchasePrice, row => row.PurchasePrice);
        Write(PriceSchema.Prices.Supply, row => row.Supply.Length > 0 ? row.Supply : null);
        Write(PriceSchema.Prices.BarcodeCopy, row => AsBarcode(row.Barcode));
        Write(PriceSchema.Prices.Units, row => row.Units);
        Write(PriceSchema.Prices.MaxPurchase, row => row.MaxPurchase);
        Write(PriceSchema.Prices.PurchaseAmount, row => row.PurchaseAmount);

        Write(PriceSchema.Prices.Code, row => Text(row.Reference?.Code));
        Write(PriceSchema.Prices.Model, row => Text(row.Reference?.Model));
        Write(PriceSchema.Prices.Sector, row => Text(row.Reference?.Sector));
        Write(PriceSchema.Prices.Group, row => Text(row.Reference?.Group));
        Write(PriceSchema.Prices.Name, row => Text(row.Reference?.Name));
        Write(PriceSchema.Prices.Article, row => Text(row.Reference?.Article));
        Write(PriceSchema.Prices.Color, row => Text(row.Reference?.Color));
        Write(PriceSchema.Prices.Season, row => Text(row.Reference?.Season));
        Write(PriceSchema.Prices.Size, row => Text(row.Reference?.Size));
        Write(PriceSchema.Prices.Subgroup, row => Text(row.Subgroup));

        // Наценки ставит пересчёт, а значения прошлой поставки оставаться не должны:
        // иначе цены посчитались бы по чужому регламенту.
        // Остальные колонки заполняет человек по ходу разбора.
        foreach (var column in new[]
                 {
                     PriceSchema.Prices.PlannedMarkup,
                     PriceSchema.Prices.MinMarkup,
                     PriceSchema.Prices.CommentKm,
                     PriceSchema.Prices.CommentBuyer,
                     PriceSchema.Prices.Decision,
                     PriceSchema.Prices.Link,
                 })
        {
            ExcelSheetOperations.ClearBlock(sheet, layout.FirstDataRow, layout.LastDataRow, headers[column]);
        }
    }

    /// <summary>
    /// Штрихкод записывается числом, если он число: в «Сводном прайсе» и в файлах
    /// согласования цен он тоже число, и текстовый штрихкод перестал бы с ними совпадать.
    /// </summary>
    private static object AsBarcode(string barcode) =>
        double.TryParse(barcode, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number)
            ? number
            : barcode;

    private static object? Text(string? value) =>
        string.IsNullOrEmpty(value) ? null : value;

    /// <summary>
    /// Каждая поставка смотрит в свою колонку листа «для цен»: первая - в «B», вторая -
    /// в «C». В книге это ссылки вида ='для цен'!$B$5, и меняется в них только буква.
    ///
    /// Проверяются все строки, а не только строки второй поставки: в прошлом файле поставок
    /// могло быть две, и строка, оставшаяся от него, смотрела бы в колонку, которой больше нет.
    /// </summary>
    private void FixSupplyReferences(object sheet, PricesLayout layout, IReadOnlyList<PriceRowValues> rows)
    {
        var columns = new[] { PriceSchema.Prices.LogisticsPercent, PriceSchema.Prices.ActualRate }
            .Select(name => layout.Headers[name])
            .ToList();

        for (var i = 0; i < rows.Count; i++)
        {
            var row = layout.FirstDataRow + i;
            var letters = ExcelColumn.ToLetters(2 + rows[i].InvoiceIndex);

            foreach (var column in columns)
            {
                var formula = ExcelSheetOperations.GetFormula(sheet, row, column);
                if (formula is null ||
                    !formula.StartsWith("=", StringComparison.Ordinal) ||
                    !TextUtils.ContainsKey(formula, PriceSchema.ForPricesSheet))
                {
                    continue;
                }

                var fixedFormula = SupplyReferenceRepair.Retarget(formula, letters);
                if (fixedFormula is not null)
                {
                    ExcelSheetOperations.SetFormula(sheet, row, column, fixedFormula);
                    _logger.Debug("Строка " + row + ": ссылка на «для цен» переведена на колонку " + letters + ".");
                }
            }
        }
    }

    /// <summary>Записывает наценки, выбранные по регламенту для каждой подгруппы.</summary>
    private void WriteMarkups(object sheet, PricesLayout layout, IReadOnlyList<PriceRowValues> rows)
    {
        var headers = layout.Headers;

        ExcelSheetOperations.SetColumnValues(
            sheet,
            layout.FirstDataRow,
            headers[PriceSchema.Prices.PlannedMarkup],
            rows.Select(row => (object?)row.Markup?.Planned).ToList());

        ExcelSheetOperations.SetColumnValues(
            sheet,
            layout.FirstDataRow,
            headers[PriceSchema.Prices.MinMarkup],
            rows.Select(row => (object?)row.Markup?.Minimum).ToList());

        _logger.Information(
            "Наценки проставлены для " + rows.Count(row => row.Markup is not null) + " из " +
            rows.Count + " строк, подгрупп в поставке: " +
            rows.Select(row => TextUtils.NormalizeKey(row.Subgroup)).Distinct().Count() + ".");
    }

    /// <summary>Сколько фотографий перенесено и что при этом не сошлось.</summary>
    private sealed record PhotoOutcome(int Copied, IReadOnlyList<ProcessingWarning> Warnings);

    /// <summary>
    /// Переносит фотографии товара из инвойса на лист «Цены» - в колонку «артикул»
    /// той строки, куда попал этот штрихкод.
    ///
    /// Старые картинки из области данных удаляются: они остались от прошлой поставки,
    /// а при копировании строки-образца ещё и размножились бы.
    /// </summary>
    private PhotoOutcome CopyInvoicePhotos(
        object applicationObject,
        IReadOnlyList<object> invoiceSheets,
        IReadOnlyList<InvoiceSheet> invoices,
        object pricesSheet,
        PricesLayout layout,
        IReadOnlyList<PriceRowValues> rows)
    {
        using var scope = new ComScope();
        var byBarcode = new Dictionary<string, object>(StringComparer.Ordinal);

        for (var index = 0; index < invoiceSheets.Count && index < invoices.Count; index++)
        {
            var whole = new SheetArea(1, ExcelConstants.LastRow, 1, ExcelConstants.LastColumn);
            var photos = ExcelPictures.ByRow(invoiceSheets[index], whole, scope);

            foreach (var line in invoices[index].Lines)
            {
                if (photos.TryGetValue(line.ExcelRow, out var shape) && !byBarcode.ContainsKey(line.Barcode))
                {
                    byBarcode[line.Barcode] = shape;
                }
            }
        }

        var area = new SheetArea(
            layout.FirstDataRow, layout.LastDataRow, layout.Columns.First, layout.Columns.Last);
        var removed = ExcelPictures.DeleteIn(pricesSheet, area);
        if (removed > 0)
        {
            _logger.Information("С листа «Цены» убрано картинок прошлой поставки: " + removed + ".");
        }

        // Фотография в инвойсе бывает одна на артикул: размеры и цвета идут строками ниже,
        // а картинка стоит только у первой. Такие строки получают фотографию своего артикула -
        // сначала того же цвета, потом любого.
        var byArticleColor = new Dictionary<string, object>(StringComparer.Ordinal);
        var byArticle = new Dictionary<string, object>(StringComparer.Ordinal);
        foreach (var priceRow in rows)
        {
            if (priceRow.Reference is { } reference &&
                TextUtils.NormalizeKey(reference.Article).Length > 0 &&
                byBarcode.TryGetValue(priceRow.Barcode, out var own))
            {
                byArticleColor.TryAdd(ArticleColorKey(reference), own);
                byArticle.TryAdd(TextUtils.NormalizeKey(reference.Article), own);
            }
        }

        var column = layout.Headers[PriceSchema.Prices.Article];
        var copied = 0;
        var fromArticle = 0;
        var missing = new List<string>();

        for (var i = 0; i < rows.Count; i++)
        {
            var row = layout.FirstDataRow + i;
            if (!byBarcode.TryGetValue(rows[i].Barcode, out var shape))
            {
                var reference = rows[i].Reference;
                if (reference is null || TextUtils.NormalizeKey(reference.Article).Length == 0 ||
                    (!byArticleColor.TryGetValue(ArticleColorKey(reference), out shape) &&
                     !byArticle.TryGetValue(TextUtils.NormalizeKey(reference.Article), out shape)))
                {
                    missing.Add(rows[i].Barcode);
                    continue;
                }

                fromArticle++;
            }

            // По инструкции фотография ставится так, чтобы артикул не был закрыт: артикул
            // стоит в ячейке сверху слева, картинка прижимается к правому нижнему углу
            // и по высоте оставляет над собой строку текста. За ячейку она не выходит.
            // Все фотографии вписываются в одну и ту же рамку: в инвойсе они разного размера,
            // и без этого одна выходила во всю строку, а другая - точкой.
            var rowHeight = ExcelSheetOperations.GetRowHeight(pricesSheet, row);
            var height = rowHeight > 2 * ArticleTextHeight ? rowHeight - ArticleTextHeight : rowHeight;
            if (ExcelPictures.CopyTo(
                    applicationObject, shape, pricesSheet, row, column,
                    PhotoWidthLimit, height, fillFrame: true, _logger, bottomRight: true))
            {
                copied++;
            }
        }

        _logger.Information(
            "Фотографий перенесено на лист «Цены»: " + copied + " из " + rows.Count +
            ", из них по артикулу (своей картинки у строки инвойса нет): " + fromArticle + ".");

        var warnings = new List<ProcessingWarning>();
        if (missing.Count > 0)
        {
            warnings.Add(new ProcessingWarning(
                "В инвойсе нет фотографии для " + missing.Count + " строк(и): " +
                string.Join(", ", missing.Take(5)) + (missing.Count > 5 ? " и другие" : "") +
                ". Картинки в этих строках листа «Цены» пустые.",
                PriceSchema.InvoiceSheet));
        }

        return new PhotoOutcome(copied, warnings);
    }

    private static string ArticleColorKey(PriceListRow reference) =>
        TextUtils.NormalizeKey(reference.Article) + "|" + TextUtils.NormalizeKey(reference.Color);

    /// <summary>
    /// Читает строки листа «Цены» такими, какие они есть: на втором этапе инвойс уже
    /// не при чём, а часть колонок аналитик успел поправить руками.
    /// </summary>
    private static IReadOnlyList<PriceRowValues> ReadPriceRows(object sheet, PricesLayout layout)
    {
        var headers = layout.Headers;
        var grid = ExcelSheetOperations.ReadBlock(
            sheet,
            layout.FirstDataRow,
            layout.LastDataRow,
            layout.Columns.First,
            layout.Columns.Last,
            withFormulas: false);

        string Text(int row, string column) => TextUtils.Normalize(grid.Text(row, headers[column]));
        double Number(int row, string column) => grid.Number(row, headers[column]) ?? 0d;

        var barcodes = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var row = layout.FirstDataRow; row <= layout.LastDataRow; row++)
        {
            var barcode = InvoiceSheetReader.NormalizeBarcode(Text(row, PriceSchema.Prices.BarcodeCopy));
            if (barcode.Length > 0)
            {
                barcodes[barcode] = barcodes.TryGetValue(barcode, out var current) ? current + 1 : 1;
            }
        }

        var rows = new List<PriceRowValues>();
        for (var row = layout.FirstDataRow; row <= layout.LastDataRow; row++)
        {
            var barcode = InvoiceSheetReader.NormalizeBarcode(Text(row, PriceSchema.Prices.BarcodeCopy));
            var sector = Text(row, PriceSchema.Prices.Sector);
            var planned = grid.Number(row, headers[PriceSchema.Prices.PlannedMarkup]);
            var minimum = grid.Number(row, headers[PriceSchema.Prices.MinMarkup]);

            rows.Add(new PriceRowValues(
                rows.Count,
                0,
                barcode,
                Text(row, PriceSchema.Prices.Supply),
                Number(row, PriceSchema.Prices.Units),
                Number(row, PriceSchema.Prices.MaxPurchase),
                Number(row, PriceSchema.Prices.PurchaseAmount),
                Number(row, PriceSchema.Prices.PurchasePrice),
                new PriceListRow(
                    barcode,
                    Text(row, PriceSchema.Prices.Code),
                    Text(row, PriceSchema.Prices.Model),
                    sector,
                    Text(row, PriceSchema.Prices.Group),
                    Text(row, PriceSchema.Prices.Name),
                    Text(row, PriceSchema.Prices.Article),
                    Text(row, PriceSchema.Prices.Color),
                    Text(row, PriceSchema.Prices.Season),
                    Text(row, PriceSchema.Prices.Size)),
                Text(row, PriceSchema.Prices.Subgroup),
                planned is null && minimum is null ? null : new MarkupRule(sector, string.Empty, planned, minimum),
                barcode.Length > 0 && barcodes[barcode] > 1));
        }

        return rows;
    }

    /// <summary>
    /// Разметка первого этапа. Колонки «Линк» и «Проверка запретов» очищаются: их заполняет
    /// второй этап, и значения прошлой поставки в них оставаться не должны.
    /// </summary>
    private void ApplyPreparedChecks(object sheet, PricesLayout layout, IReadOnlyList<PriceRowCheck> checks)
    {
        var headers = layout.Headers;

        foreach (var column in new[] { PriceSchema.Prices.Link, PriceSchema.Prices.BanCheck })
        {
            ExcelSheetOperations.ClearBlock(sheet, layout.FirstDataRow, layout.LastDataRow, headers[column]);
        }

        var marked = 0;
        for (var i = 0; i < checks.Count; i++)
        {
            var row = layout.FirstDataRow + i;
            var highlight = checks[i].Highlight.ToHashSet(StringComparer.Ordinal);

            foreach (var column in MarkableColumns(headers))
            {
                var mark = highlight.Contains(column) ? RowMark.Attention : RowMark.None;
                ExcelSheetOperations.SetRowMark(sheet, row, headers[column], mark);
                if (mark == RowMark.Attention)
                {
                    marked++;
                }
            }
        }

        _logger.Information("Подсвечено ячеек: " + marked + ".");
    }

    /// <summary>
    /// Формулы под таблицей: итоговая строка и всё, что ниже. Их диапазоны растягиваются
    /// на фактические строки данных - новые строки добавляются под последней, за границей
    /// диапазона, и сам Excel его не растягивает. Если в итоговой строке у «Ед.»
    /// и «Сумма_закупа» формулы нет, пишутся посчитанные суммы.
    /// </summary>
    private void WriteTotals(object sheet, PricesLayout layout, IReadOnlyList<PriceRowValues> rows)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(sheet);
        var lastRow = Math.Max(layout.TotalsRow, Math.Min(bounds.LastRow, layout.TotalsRow + BelowTableRows));
        var grid = ExcelSheetOperations.ReadBlock(
            sheet, layout.TotalsRow, lastRow, layout.Columns.First, layout.Columns.Last, withFormulas: true);

        var extended = 0;
        for (var row = layout.TotalsRow; row <= lastRow; row++)
        {
            for (var column = layout.Columns.First; column <= layout.Columns.Last; column++)
            {
                var repaired = TotalsFormulaRepair.ExtendBelow(
                    grid.Formula(row, column), layout.FirstDataRow, layout.LastDataRow, layout.TotalsRow);
                if (repaired is null)
                {
                    continue;
                }

                ExcelSheetOperations.SetFormula(sheet, row, column, repaired);
                extended++;
                _logger.Debug(new CellRef(row, column) + ": формула под таблицей растянута - " + repaired + ".");
            }
        }

        _logger.Information("Формул под таблицей «Цены» растянуто на все строки: " + extended + ".");

        void Total(string column, double value)
        {
            var index = layout.Headers[column];
            if (!grid.HasFormula(layout.TotalsRow, index))
            {
                ExcelSheetOperations.SetValue(sheet, layout.TotalsRow, index, value);
            }
        }

        Total(PriceSchema.Prices.Units, rows.Sum(row => row.Units));
        Total(PriceSchema.Prices.PurchaseAmount, rows.Sum(row => row.PurchaseAmount));
    }

    /// <summary>
    /// Все строки данных «Цены» получают формулы, оформление и высоту первой строки.
    ///
    /// Новые строки и так копируются с первой, но строки, которые уже стояли в заготовке,
    /// могли остаться пустыми: после прошлого распреда одна строка заполнена, а вторая - нет,
    /// и без этого шага во второй строке не было бы ни одной формулы. Формула переносится
    /// в относительной записи, как при протягивании, - каждая строка считает свою.
    /// Колонки, где у первой строки значение, а не формула, не трогаются: их заполняет
    /// программа или человек.
    /// </summary>
    private void NormalizeDataRows(object application, object sheet, PricesLayout layout)
    {
        var first = layout.FirstDataRow;
        var height = Math.Max(ExcelSheetOperations.GetRowHeight(sheet, first), MinimumPhotoRowHeight);
        ExcelSheetOperations.SetRowsHeight(sheet, first, layout.LastDataRow, height);

        if (layout.DataRowCount < 2)
        {
            return;
        }

        ExcelSheetOperations.CopyRowFormats(application, sheet, first, first + 1, layout.LastDataRow, layout.Columns);

        // Дописываются только ячейки, где формулы образца нет, - подряд идущие одним куском.
        // Формулы именно копируются из первой строки, а не вписываются: вписанную формулу
        // Excel сразу считает, и СРЗНАЧЕСЛИ по стотысячному «прайсу» занимает секунды
        // на ячейку; при копировании он этого не делает.
        var template = ExcelSheetOperations.ReadRowFormulasR1C1(sheet, first, layout.Columns.First, layout.Columns.Last);
        var filled = 0;
        for (var row = first + 1; row <= layout.LastDataRow; row++)
        {
            var current = ExcelSheetOperations.ReadRowFormulasR1C1(sheet, row, layout.Columns.First, layout.Columns.Last);
            var start = -1;
            for (var i = 0; i <= template.Length; i++)
            {
                var differs = i < template.Length &&
                              template[i] is not null &&
                              !string.Equals(template[i], current[i], StringComparison.Ordinal);
                if (differs && start < 0)
                {
                    start = i;
                }
                else if (!differs && start >= 0)
                {
                    var firstColumn = layout.Columns.First + start;
                    ExcelSheetOperations.CopyRange(
                        application, sheet, first, firstColumn, first, firstColumn + (i - start) - 1, row, firstColumn);
                    filled += i - start;
                    start = -1;
                }
            }
        }

        _logger.Information(
            "Строки «Цены» " + (first + 1) + "-" + layout.LastDataRow + " приведены к первой: дописано формул " +
            filled + ", высота строк " + height.ToString("0.##", CultureInfo.InvariantCulture) + ".");
    }

    // ================= Проверки =================

    private static IReadOnlyList<PriceRowState> ReadRowStates(
        object sheet,
        PricesLayout layout,
        IReadOnlyList<PriceRowValues> rows)
    {
        var grid = ExcelSheetOperations.ReadBlock(
            sheet,
            layout.FirstDataRow,
            layout.LastDataRow,
            layout.Columns.First,
            layout.Columns.Last,
            withFormulas: false);

        var headers = layout.Headers;
        var states = new List<PriceRowState>(rows.Count);
        var gradeColumns = PriceSchema.Prices.AllGrades
            .Where(grade => headers.TryGet(grade, out _))
            .ToList();

        for (var i = 0; i < rows.Count; i++)
        {
            var row = layout.FirstDataRow + i;

            var grades = gradeColumns
                .Where(grade => grid.Number(row, headers[grade]) is 1d)
                .ToList();

            states.Add(new PriceRowState(
                i,
                TextUtils.Normalize(grid.Text(row, headers[PriceSchema.Prices.VspR])),
                grid.Value(row, headers[PriceSchema.Prices.StoreSalesPeriod]),
                grid.Value(row, headers[PriceSchema.Prices.IslandSalesPeriod]),
                grid.Value(row, headers[PriceSchema.Prices.WallAnalogue]),
                grid.Value(row, headers[PriceSchema.Prices.GoogleAnalogue]),
                grid.Value(row, headers[PriceSchema.Prices.PriceCheck]),
                grades,
                headers.TryGet(PriceSchema.Prices.Marketplace, out var marketplace)
                    ? grid.Value(row, marketplace)
                    : null,
                gradeColumns));
        }

        return states;
    }

    private static IReadOnlyList<PriceRowCheck> RunChecks(
        IReadOnlyList<PriceRowValues> rows,
        IReadOnlyList<PriceRowState> states,
        IReadOnlyDictionary<string, LinkEntry> links)
    {
        var checks = new List<PriceRowCheck>(rows.Count);
        for (var i = 0; i < rows.Count; i++)
        {
            links.TryGetValue(TextUtils.NormalizeKey(states[i].Acr), out var link);
            checks.Add(PriceChecks.Check(rows[i], states[i], link));
        }

        return checks;
    }

    /// <summary>
    /// Привязывает замечания листа «Цены» к ячейкам. Строка известна по номеру строки
    /// проверки, колонка - по первой подсвеченной: именно её и нужно увидеть, открыв книгу.
    /// Проверки считаются по строкам в памяти и адресов не знают - поэтому адреса
    /// проставляются здесь, где разметка листа уже прочитана.
    /// </summary>
    private static IReadOnlyList<PriceRowCheck> Locate(
        object sheet, PricesLayout layout, IReadOnlyList<PriceRowCheck> checks)
    {
        var sheetName = ExcelSheetOperations.GetSheetName(sheet);
        var fallback = layout.Headers[PriceSchema.Prices.BarcodeCopy];
        var located = new List<PriceRowCheck>(checks.Count);

        foreach (var check in checks)
        {
            if (check.Warnings.Count == 0)
            {
                located.Add(check);
                continue;
            }

            var column = fallback;
            foreach (var name in check.Highlight)
            {
                if (layout.Headers.TryGet(name, out var found))
                {
                    column = found;
                    break;
                }
            }

            var cell = new CellRef(layout.FirstDataRow + check.Index, column).ToString();
            located.Add(check with
            {
                Warnings = check.Warnings
                    .Select(warning => warning with { Sheet = sheetName, Cell = cell })
                    .ToList(),
            });
        }

        return located;
    }

    /// <summary>
    /// Записывает «Линк» и «Проверка запретов» и красит проблемные ячейки.
    ///
    /// Пометка снимается со всех ячеек, которые программа может пометить, и ставится
    /// заново - иначе вчерашняя пометка осталась бы на строке, где всё уже сошлось.
    /// Снимается только своя заливка: цвет, поставленный аналитиком, сохраняется.
    /// </summary>
    private void ApplyChecks(object sheet, PricesLayout layout, IReadOnlyList<PriceRowCheck> checks)
    {
        var headers = layout.Headers;

        ExcelSheetOperations.SetColumnValues(
            sheet,
            layout.FirstDataRow,
            headers[PriceSchema.Prices.Link],
            checks.Select(check => (object?)check.LinkValue).ToList());

        ExcelSheetOperations.SetColumnValues(
            sheet,
            layout.FirstDataRow,
            headers[PriceSchema.Prices.BanCheck],
            checks.Select(check => check.BanCheck.Length > 0 ? (object?)check.BanCheck : null).ToList());

        var markable = MarkableColumns(headers);
        var marked = 0;

        for (var i = 0; i < checks.Count; i++)
        {
            var row = layout.FirstDataRow + i;
            var highlight = checks[i].Highlight.ToHashSet(StringComparer.Ordinal);

            foreach (var column in markable)
            {
                var mark = highlight.Contains(column) ? RowMark.Attention : RowMark.None;
                ExcelSheetOperations.SetRowMark(sheet, row, headers[column], mark);
                if (mark == RowMark.Attention)
                {
                    marked++;
                }
            }
        }

        _logger.Information("Подсвечено ячеек: " + marked + ".");
    }

    /// <summary>
    /// Одинаковые штрихкоды выделяются цветом в «Копия ШК»: у каждой группы повторов свой,
    /// так пары видно сразу, в том числе когда повторы пришли из разных инвойсов.
    /// Вызывается после разметки проверок: та уже сняла прошлые цвета, а розовая пометка
    /// «нет в прайсе» у повторов общая - об этом всё равно пишется замечание.
    /// </summary>
    private void MarkDuplicateBarcodes(object sheet, PricesLayout layout, IReadOnlyList<PriceRowValues> rows)
    {
        var column = layout.Headers[PriceSchema.Prices.BarcodeCopy];
        var groups = DuplicateBarcodeGroups.Assign(rows.Select(row => row.Barcode).ToList());

        for (var i = 0; i < groups.Count; i++)
        {
            if (groups[i] is { } group)
            {
                ExcelSheetOperations.SetFill(sheet, layout.FirstDataRow + i, column, DuplicateBarcodeGroups.ColorOf(group));
            }
        }

        var count = groups.Where(group => group is not null).Distinct().Count();
        if (count > 0)
        {
            _logger.Information("Одинаковые штрихкоды выделены цветом в «Копия ШК»: групп " + count + ".");
        }
    }

    /// <summary>Колонки, которые программа может пометить. Необязательных на листе может не оказаться.</summary>
    private static IReadOnlyList<string> MarkableColumns(HeaderMap headers) => new[]
    {
        PriceSchema.Prices.BarcodeCopy,
        PriceSchema.Prices.PurchasePrice,
        PriceSchema.Prices.PlannedMarkup,
        PriceSchema.Prices.StoreSalesPeriod,
        PriceSchema.Prices.GoogleAnalogue,
        PriceSchema.Prices.PriceCheck,
        PriceSchema.Prices.Link,
        PriceSchema.Prices.Marketplace,
    }.Concat(PriceSchema.Prices.AllGrades)
        .Where(column => headers.TryGet(column, out _))
        .ToList();

    private static void Report(IProgress<ProcessingStage>? progress, string message, int percent) =>
        progress?.Report(new ProcessingStage(message, percent));
}
