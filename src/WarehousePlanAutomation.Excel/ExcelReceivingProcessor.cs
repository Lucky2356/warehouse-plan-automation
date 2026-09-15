using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>Какой из двух этапов приемки выполняется.</summary>
public enum ReceivingStage
{
    /// <summary>
    /// Подготовка: выгрузки разобраны, «Приходы» покрашены, листы поставок и «итог» собраны.
    /// После неё аналитик по секторам решает, что написать в «Допоставить».
    /// </summary>
    Prepare,

    /// <summary>
    /// Адреса: «Допоставить» решено, программа собирает «адреса» и «тары» и подтягивает
    /// на «итог» место хранения и тару.
    /// </summary>
    Addresses,
}

/// <summary>
/// Книга «Приемка на хранилище». Работа разделена на два этапа, потому что между ними
/// человек делает то, что за него не решить: по каждому сектору смотрит остатки, продажи,
/// прогноз и даты стенок и пишет, сколько допоставить.
///
/// Исходный файл никогда не изменяется: все действия выполняются над копией.
/// По сети программа не ходит: колонки со стенками на «итоге» считает сам Excel
/// по сохранённым в книге значениям.
/// </summary>
public sealed class ExcelReceivingProcessor : IWorkbookProcessor
{
    private const int HeaderScanRows = 15;
    private const int ChunkRows = 50000;

    /// <summary>
    /// Цвета «Приходов» - те же, что в шаблоне аналитика (Interior.Color, порядок BGR):
    /// розовый - маркетплейс, зелёный - собрана, голубой - не собрана.
    /// </summary>
    private const int MarketplaceFill = 0xCCCCFF;

    private const int CollectedFill = 0xCCFFCC;
    private const int NotCollectedFill = 0xEED7BD;

    private static readonly int[] OwnFills = { MarketplaceFill, CollectedFill, NotCollectedFill };

    private readonly IAppLogger _logger;
    private readonly ReceivingStage _stage;
    private readonly Func<DateTime> _nowProvider;

    public ExcelReceivingProcessor(
        IAppLogger logger,
        ReceivingStage stage = ReceivingStage.Prepare,
        Func<DateTime>? nowProvider = null)
    {
        _logger = logger;
        _stage = stage;
        _nowProvider = nowProvider ?? (() => DateTime.Now);
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
            _stage == ReceivingStage.Prepare ? OutputFileName.ReceivingMark : OutputFileName.AddressesMark);
        _logger.Information(
            (_stage == ReceivingStage.Prepare ? "Начало подготовки приемки." : "Начало расстановки адресов приемки.") +
            " Исходный файл: " + sourceFilePath);
        Report(progress, "Создание копии файла", 3);

        File.Copy(sourceFilePath, outputPath);
        _logger.Information("Создана копия: " + outputPath);

        try
        {
            var result = ProcessCopy(outputPath, progress, cancellationToken);
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

            var outcome = _stage == ReceivingStage.Prepare
                ? Prepare(applicationObject, (object)workbook, path, progress, cancellationToken, scope)
                : PlaceAddresses(applicationObject, (object)workbook, path, progress, cancellationToken, scope);

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

    private ProcessingResult Prepare(
        object applicationObject,
        object workbook,
        string path,
        IProgress<ProcessingStage>? progress,
        CancellationToken cancellationToken,
        ComScope scope)
    {
        dynamic application = applicationObject;

        var problems = new List<string>();
        var stockSheet = Required(workbook, ReceivingSchema.StockSheet, scope, problems);
        var storageStockSheet = Required(workbook, ReceivingSchema.StorageStockSheet, scope, problems);
        var reservesSheet = Required(workbook, ReceivingSchema.ReservesSheet, scope, problems);
        var storageSheet = Required(workbook, ReceivingSchema.StorageSheet, scope, problems);
        var marketplaceSheet = Required(workbook, ReceivingSchema.MarketplaceSheet, scope, problems);
        var incomingSheet = Required(workbook, ReceivingSchema.IncomingSheet, scope, problems);
        var summarySheet = Required(workbook, ReceivingSchema.SummarySheet, scope, problems);
        var collectedSheet = Required(workbook, ReceivingSchema.CollectedSheet, scope, problems);
        var notCollectedSheet = Required(workbook, ReceivingSchema.NotCollectedSheet, scope, problems);
        var marketplaceSuppliesSheet = Required(workbook, ReceivingSchema.MarketplaceSuppliesSheet, scope, problems);
        var warehouseSheet = Required(workbook, ReceivingSchema.WarehouseSheet, scope, problems);
        var removedSheet = Required(workbook, ReceivingSchema.RemovedFromPlanSheet, scope, problems);
        var planSheet = Required(workbook, ReceivingSchema.PlanSheet, scope, problems);
        var notAcceptedSheet = Required(workbook, ReceivingSchema.NotAcceptedSheet, scope, problems);
        var agreementSheet = Optional(workbook, ReceivingSchema.AgreementSheet, scope, problems);

        if (problems.Count > 0)
        {
            throw new WorkbookValidationException(problems);
        }

        var warnings = new List<ProcessingWarning>();
        var listSeparator = ExcelSheetOperations.GetListSeparator(applicationObject);
        var today = _nowProvider();

        Report(progress, "Снятие фильтров на всех листах", 6);
        ShowAllRowsEverywhere(workbook, scope);

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Лист «Т.Остатки»: брак", 10);
        var storageSums = CleanStorageStock(storageStockSheet!, listSeparator, warnings);

        Report(progress, "Лист «Остатки»", 16);
        var stock = CleanStock(stockSheet!, storageSums, listSeparator, warnings);

        Report(progress, "Лист «Резервы»", 24);
        CleanReserves(reservesSheet!, today.Year, listSeparator, warnings);

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Чтение листа «Склад»", 30);
        var warehouse = ReadWarehouse(warehouseSheet!);

        Report(progress, "Листы «А2, А3» и «МП»", 40);
        var storage = StorageSelection.Select(warehouse, ReceivingSchema.Warehouse.StorageTypeStorage);
        var marketplace = StorageSelection.Select(warehouse, ReceivingSchema.Warehouse.StorageTypeMarketplace);
        var summaryColumns = ReadSummaryColumns(summarySheet!);
        WriteStorage(
            applicationObject, storageSheet!, storage, ReceivingSchema.Warehouse.StorageTypeStorage, summaryColumns, warnings);
        WriteStorage(
            applicationObject, marketplaceSheet!, marketplace, ReceivingSchema.Warehouse.StorageTypeMarketplace, summaryColumns, warnings);

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Лист «Приходы»", 50);
        var incoming = SortIncoming(incomingSheet!, removedSheet!, planSheet!, listSeparator, warnings);

        Report(progress, "Листы поставок", 60);
        var notAccepted = ReadNotAccepted(notAcceptedSheet!);
        var collected = WriteSupplies(applicationObject, collectedSheet!, notAccepted, incoming.Collected, warnings);
        var notCollected = WriteSupplies(applicationObject, notCollectedSheet!, notAccepted, incoming.NotCollected, warnings);
        var marketplaceSupplies = WriteSupplies(
            applicationObject, marketplaceSuppliesSheet!, notAccepted, incoming.Marketplace, warnings);

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Лист «итог»", 72);
        IReadOnlyDictionary<string, double> answers;
        if (agreementSheet is null)
        {
            answers = new Dictionary<string, double>(StringComparer.Ordinal);
            warnings.Add(new ProcessingWarning(
                "Листа «" + ReceivingSchema.AgreementSheet + "» в книге нет - «Количество МП» на «итоге» " +
                "осталось таким, как его считает формула.",
                ReceivingSchema.SummarySheet));
        }
        else
        {
            answers = ReadAnswers(agreementSheet);
        }

        var sources = new SummarySources(
            QuantitiesByCode(storage.Rows),
            QuantitiesByCode(marketplace.Rows),
            collected.DeviationByCode,
            notCollected.DeviationByCode,
            marketplaceSupplies.DeviationByCode);
        var summary = WriteSummary(
            applicationObject, summarySheet!, stock, sources, answers, MonthCalendar.WeekNumber(today), warnings);

        Report(progress, "Пересчёт формул", 86);
        application.Calculation = ExcelConstants.XlCalculationAutomatic;
        application.CalculateFull();

        return new ProcessingResult(
            path,
            new[]
            {
                new ProcessingCounter("строк «итога»", summary.Rows),
                new ProcessingCounter("адресов «А2, А3»", storage.Rows.Count),
                new ProcessingCounter("адресов «МП»", marketplace.Rows.Count),
                new ProcessingCounter("поставок собрано", collected.Supplies),
                new ProcessingCounter("поставок не собрано", notCollected.Supplies),
                new ProcessingCounter("поставок МП", marketplaceSupplies.Supplies),
                new ProcessingCounter("не разобрано в «Приходах»", incoming.Unresolved, incoming.Unresolved > 0),
            },
            warnings);
    }

    /// <summary>
    /// «Т.Остатки»: строки брака удаляются, остальное складывается по коду - это то,
    /// что вручную подтягивает СУММЕСЛИМН на лист «Остатки».
    /// </summary>
    private Dictionary<string, double> CleanStorageStock(
        object sheet, string listSeparator, List<ProcessingWarning> warnings)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.StorageStock.Specs);
        var codeColumn = table.Headers[ReceivingSchema.StorageStock.Code];
        var articleColumn = table.Headers[ReceivingSchema.StorageStock.Article];
        var remainderColumn = table.Headers[ReceivingSchema.StorageStock.Remainder];
        var first = table.Headers.HeaderRow + 1;
        var last = LastRow(sheet, table.Bounds, codeColumn, articleColumn);

        if (last < first)
        {
            warnings.Add(new ProcessingWarning(
                "Лист «" + name + "» пуст: остаток на хранилище везде будет нулевым.", name));
            return new Dictionary<string, double>(StringComparer.Ordinal);
        }

        var grid = ExcelSheetOperations.ReadBlock(
            sheet, first, last, 1, Math.Max(codeColumn, Math.Max(articleColumn, remainderColumn)), withFormulas: false);

        var defects = new List<int>();
        var kept = new List<(object?, object?)>();
        var fromArticle = 0;

        for (var row = first; row <= last; row++)
        {
            object? code = grid.Value(row, codeColumn);
            if (CellError.IsError(code) || TextUtils.Normalize(TextUtils.CellToString(code)).Length == 0)
            {
                var article = grid.Value(row, articleColumn);
                if (TextUtils.Normalize(TextUtils.CellToString(article)).Length == 0)
                {
                    continue;
                }

                code = ReceivingStockRules.CodeFromArticle(article);
                fromArticle++;
            }

            if (ReceivingStockRules.IsDefect(code))
            {
                defects.Add(row);
                continue;
            }

            kept.Add((code, grid.Value(row, remainderColumn)));
        }

        ExcelSheetOperations.DeleteRows(sheet, defects, listSeparator);
        _logger.Information("«" + name + "»: удалено строк брака " + defects.Count + ", осталось " + kept.Count + ".");

        if (fromArticle > 0)
        {
            warnings.Add(new ProcessingWarning(
                "На листе «" + name + "» у " + fromArticle + " строк не заполнен «Код»: формула не дотянута " +
                "до новых строк выгрузки. Для расчёта код взят из последних шести знаков «Артикула» - " +
                "так же, как в самой формуле.",
                name));
        }

        return ReceivingStockRules.SumByCode(kept);
    }

    /// <summary>
    /// «Остатки»: удаляются строки «Denny goods» = «все» и все колонки, кроме одиннадцати
    /// нужных, а «Остаток хранилище» пересчитывается по «Т.Остаткам» и пишется значениями.
    /// </summary>
    private StockRows CleanStock(
        object sheet,
        IReadOnlyDictionary<string, double> storageSums,
        string listSeparator,
        List<ProcessingWarning> warnings)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.Stock.Specs);
        var headerRow = table.Headers.HeaderRow;
        var first = headerRow + 1;
        var last = LastRow(
            sheet, table.Bounds, table.Headers[ReceivingSchema.Stock.Code], table.Headers[ReceivingSchema.Stock.Article]);

        var denyColumn = FindColumn(sheet, headerRow, table.LastHeaderColumn, ReceivingSchema.Stock.DenyGoodsAliases);
        if (denyColumn is { } deny && last >= first)
        {
            var denyGrid = ExcelSheetOperations.ReadBlock(sheet, first, last, deny, deny, withFormulas: false);
            var denied = Enumerable.Range(first, last - first + 1)
                .Where(row => ReceivingStockRules.IsDeniedForAll(denyGrid.Value(row, deny)))
                .ToList();
            ExcelSheetOperations.DeleteRows(sheet, denied, listSeparator);
            _logger.Information("«" + name + "»: удалено строк «Denny goods» = «все»: " + denied.Count + ".");
        }

        var keep = ReceivingSchema.Stock.Specs.Select(spec => table.Headers[spec.DisplayName]).ToHashSet();
        var extra = Enumerable.Range(1, table.LastHeaderColumn).Where(column => !keep.Contains(column)).ToList();
        foreach (var (start, count) in DescendingRuns(extra))
        {
            ExcelSheetOperations.DeleteColumns(sheet, start, count);
        }

        if (extra.Count > 0)
        {
            _logger.Information("«" + name + "»: удалено колонок " + extra.Count + ", осталось " + keep.Count + ".");
        }

        table = Headers(sheet, ReceivingSchema.Stock.Specs);
        var codeColumn = table.Headers[ReceivingSchema.Stock.Code];
        first = table.Headers.HeaderRow + 1;
        last = LastRow(sheet, table.Bounds, codeColumn, table.Headers[ReceivingSchema.Stock.Article]);

        if (last < first)
        {
            warnings.Add(new ProcessingWarning("Лист «" + name + "» пуст: «итог» будет пустым.", name));
            return new StockRows(Array.Empty<object?[]>(), Array.Empty<object?>());
        }

        var lastColumn = ReceivingSchema.Stock.Specs.Max(spec => table.Headers[spec.DisplayName]);
        var grid = ExcelSheetOperations.ReadBlock(sheet, first, last, 1, lastColumn, withFormulas: false);
        var remainderColumn = table.Headers[ReceivingSchema.Stock.StorageRemainder];

        var remainders = new List<object?>(last - first + 1);
        var rows = new List<object?[]>(last - first + 1);
        var codes = new List<object?>(last - first + 1);

        for (var row = first; row <= last; row++)
        {
            var code = grid.Value(row, codeColumn);
            var key = ReceivingStockRules.CodeKey(code);
            var remainder = key.Length > 0 && storageSums.TryGetValue(key, out var sum) ? sum : 0d;

            remainders.Add(remainder);
            codes.Add(code);
            rows.Add(ReceivingSchema.Summary.FromStock
                .Select(column => column == ReceivingSchema.Stock.StorageRemainder
                    ? remainder
                    : grid.Value(row, table.Headers[column]))
                .ToArray());
        }

        ExcelSheetOperations.SetColumnValues(sheet, first, remainderColumn, remainders);
        _logger.Information(
            "«" + name + "»: строк " + rows.Count + ", «Остаток хранилище» больше нуля у " +
            remainders.Count(value => value is double number && number > 0) + ".");

        return new StockRows(rows, codes);
    }

    /// <summary>«Резервы»: удаляются резервы прошлых лет и всё, что не «РезервОтгрузка».</summary>
    private void CleanReserves(object sheet, int year, string listSeparator, List<ProcessingWarning> warnings)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.Reserves.Specs);
        var dateColumn = table.Headers[ReceivingSchema.Reserves.Date];
        var typeColumn = table.Headers[ReceivingSchema.Reserves.Type];
        var codeColumn = table.Headers[ReceivingSchema.Reserves.Code];
        var first = table.Headers.HeaderRow + 1;
        var last = LastRow(sheet, table.Bounds, codeColumn, dateColumn, typeColumn);

        if (last < first)
        {
            return;
        }

        var grid = ExcelSheetOperations.ReadBlock(
            sheet, first, last, 1, Math.Max(dateColumn, Math.Max(typeColumn, codeColumn)), withFormulas: false);

        var remove = new List<int>();
        var outdated = 0;
        var otherTypes = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
        var otherCount = 0;

        for (var row = first; row <= last; row++)
        {
            var date = grid.Value(row, dateColumn);
            var type = grid.Value(row, typeColumn);
            if (Blank(date) && Blank(type) && Blank(grid.Value(row, codeColumn)))
            {
                continue;
            }

            if (ReceivingStockRules.IsOutdatedReserve(date, year))
            {
                outdated++;
                remove.Add(row);
                continue;
            }

            if (!ReceivingStockRules.IsShipmentReserve(type))
            {
                otherCount++;
                otherTypes.Add(TextUtils.Normalize(TextUtils.CellToString(type)) is { Length: > 0 } text ? text : "пусто");
                remove.Add(row);
            }
        }

        ExcelSheetOperations.DeleteRows(sheet, remove, listSeparator);
        _logger.Information("«" + name + "»: удалено резервов прошлых лет " + outdated + ", другого типа " + otherCount + ".");

        if (otherCount > 0)
        {
            warnings.Add(new ProcessingWarning(
                "На листе «" + name + "» были резервы не «" + ReceivingSchema.Reserves.ShipmentReserve + "» (" +
                string.Join(", ", otherTypes) + "): " + otherCount + " шт. По инструкции на листе остаётся " +
                "только этот тип - строки удалены, в «Резервы» на «итоге» они не попадут.",
                name));
        }
    }

    /// <summary>Строки «Склада» двух нужных типов хранения. Лист большой - читается порциями.</summary>
    private List<WarehouseRow> ReadWarehouse(object sheet)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.Warehouse.Specs);
        var headers = table.Headers;
        var typeColumn = headers[ReceivingSchema.Warehouse.StorageType];
        var first = headers.HeaderRow + 1;
        var last = LastRow(
            sheet, table.Bounds, typeColumn, headers[ReceivingSchema.Warehouse.Code], headers[ReceivingSchema.Warehouse.Address]);

        var columns = ReceivingSchema.Warehouse.Specs.Select(spec => headers[spec.DisplayName]).ToList();
        var firstColumn = columns.Min();
        var lastColumn = columns.Max();
        var wanted = new[]
        {
            TextUtils.NormalizeKey(ReceivingSchema.Warehouse.StorageTypeStorage),
            TextUtils.NormalizeKey(ReceivingSchema.Warehouse.StorageTypeMarketplace),
        };
        var sources = ReceivingSchema.Storage.FromWarehouse
            .Select(pair => pair.Source is null ? (int?)null : headers[pair.Source])
            .ToList();

        var rows = new List<WarehouseRow>();
        for (var start = first; start <= last; start += ChunkRows)
        {
            var end = Math.Min(start + ChunkRows - 1, last);
            var grid = ExcelSheetOperations.ReadBlock(sheet, start, end, firstColumn, lastColumn, withFormulas: false);

            for (var row = start; row <= end; row++)
            {
                var type = grid.Value(row, typeColumn);
                if (Array.IndexOf(wanted, TextUtils.NormalizeKey(TextUtils.CellToString(type))) < 0)
                {
                    continue;
                }

                rows.Add(new WarehouseRow(
                    row,
                    type,
                    grid.Value(row, headers[ReceivingSchema.Warehouse.Container]),
                    grid.Value(row, headers[ReceivingSchema.Warehouse.Code]),
                    grid.Value(row, headers[ReceivingSchema.Warehouse.Quantity]),
                    sources.Select(column => column is { } c ? grid.Value(row, c) : null).ToArray()));
            }
        }

        _logger.Information("«" + name + "»: строк " + Math.Max(last - first + 1, 0) + ", нужных типов хранения " + rows.Count + ".");
        return rows;
    }

    /// <summary>
    /// «А2, А3» и «МП» со «Склада». Колонки до «Цвет» заполняются значениями по названиям,
    /// формулы первой строки протягиваются на все строки. Если в «В приемку» формулы нет -
    /// она теряется, когда вставка со «Склада» съезжает на колонку, - программа вписывает
    /// её сама: без неё «Учитывать» не набирает количество под «Допоставить».
    /// </summary>
    private void WriteStorage(
        object applicationObject,
        object sheet,
        StorageSelectionResult selection,
        string storageType,
        SummaryColumns summary,
        List<ProcessingWarning> warnings)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.Storage.Specs);
        var headers = table.Headers;
        var first = headers.HeaderRow + 1;
        var width = new ColumnRange(1, table.LastHeaderColumn);
        var toReceive = headers[ReceivingSchema.Storage.ToReceive];
        var oldLast = LastRow(
            sheet,
            table.Bounds,
            headers[ReceivingSchema.Storage.Address],
            headers[ReceivingSchema.Storage.Code],
            headers[ReceivingSchema.Storage.Counted],
            toReceive);

        var template = ExcelSheetOperations.ReadBlock(sheet, first, first, 1, width.Last, withFormulas: true);
        var hasToReceiveFormula = template.HasFormula(first, toReceive);
        var count = selection.Rows.Count;
        var lastRow = Math.Max(first, first + count - 1);

        var codeColumn = headers[ReceivingSchema.Storage.Code];
        if (!hasToReceiveFormula)
        {
            // Формула ставится в строку-образец до протягивания - и расходится вместе с остальными.
            ExcelSheetOperations.SetFormula(
                sheet,
                first,
                toReceive,
                ReceivingFormulas.ToReceive(summary.Sheet, summary.CodeColumn, summary.RestockColumn, codeColumn, first));
        }

        ExcelSheetOperations.ClearRange(sheet, first + 1, oldLast, width.First, width.Last);
        FillDown(applicationObject, sheet, first, first + count - 1, width);

        for (var i = 0; i < ReceivingSchema.Storage.FromWarehouse.Count; i++)
        {
            var column = headers[ReceivingSchema.Storage.FromWarehouse[i].Target];
            if (count == 0)
            {
                ExcelSheetOperations.ClearBlock(sheet, first, first, column);
                continue;
            }

            var index = i;
            WriteColumn(sheet, first, column, selection.Rows.Select(row => row.Values[index]).ToList());
        }

        if (!hasToReceiveFormula)
        {
            warnings.Add(ToReceiveRestored(
                name, first, toReceive, TextUtils.Normalize(template.Text(first, toReceive)), summary, codeColumn));
        }

        if (selection.BlankContainers > 0)
        {
            warnings.Add(new ProcessingWarning(
                "На «Складе» у " + (selection.BlankContainers + 1) + " строк типа «" + storageType +
                "» пустая «Тара». Удаление дубликатов по таре оставило из них одну - так же, " +
                "как это делает Excel. Проверьте эти адреса на «Складе».",
                name));
        }

        ReapplyAutoFilter(sheet, lastRow);
        _logger.Information(
            "«" + name + "»: строк «" + storageType + "» на «Складе» " + selection.Matched +
            ", повторов тары убрано " + selection.DuplicateContainers + ", записано " + count + ".");

        if (count == 0)
        {
            warnings.Add(new ProcessingWarning(
                "На «Складе» нет ни одной строки с типом хранения «" + storageType + "»: лист «" + name + "» пуст.",
                name));
        }
    }

    /// <summary>«Приходы»: лишнее удаляется, остальное красится по тому, что это за поставка.</summary>
    private IncomingOutcome SortIncoming(
        object sheet,
        object removedSheet,
        object planSheet,
        string listSeparator,
        List<ProcessingWarning> warnings)
    {
        var removedNumbers = SupplyNumbers.ExtractAll(AllTexts(removedSheet));
        var plannedNumbers = ReadPlanNumbers(planSheet);

        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.Incoming.Specs);
        var numberColumn = table.Headers[ReceivingSchema.Incoming.Number];
        var statusColumn = table.Headers[ReceivingSchema.Incoming.Status];
        var differenceColumn = table.Headers[ReceivingSchema.Incoming.Difference];
        var first = table.Headers.HeaderRow + 1;
        var last = LastRow(sheet, table.Bounds, numberColumn, statusColumn);

        var empty = new HashSet<string>(StringComparer.Ordinal);
        if (last < first)
        {
            warnings.Add(new ProcessingWarning("Лист «" + name + "» пуст: листы поставок будут пустыми.", name));
            return new IncomingOutcome(empty, empty, empty, 0);
        }

        var grid = ExcelSheetOperations.ReadBlock(
            sheet, first, last, 1, Math.Max(numberColumn, Math.Max(statusColumn, differenceColumn)), withFormulas: false);
        var rows = Enumerable.Range(0, last - first + 1)
            .Select(i => new IncomingRow(
                i,
                grid.NormalizedText(first + i, numberColumn),
                grid.Value(first + i, statusColumn),
                grid.Value(first + i, differenceColumn)))
            .ToList();

        var decisions = IncomingSupplyRules.Decide(rows, removedNumbers, plannedNumbers);
        var removed = decisions.Where(d => d.Kind == IncomingKind.Removed).ToList();
        ExcelSheetOperations.DeleteRows(sheet, removed.Select(d => first + d.Index), listSeparator);

        var collected = new HashSet<string>(StringComparer.Ordinal);
        var notCollected = new HashSet<string>(StringComparer.Ordinal);
        var marketplace = new HashSet<string>(StringComparer.Ordinal);
        var unresolved = new List<(string Number, string Cell, string Note)>();
        var position = 0;

        foreach (var decision in decisions)
        {
            if (decision.Kind == IncomingKind.Removed)
            {
                continue;
            }

            var row = first + position++;
            var number = rows[decision.Index].Number;
            var key = SupplyNumbers.Key(number);
            var cell = new CellRef(row, numberColumn).ToString();

            switch (decision.Kind)
            {
                case IncomingKind.Marketplace:
                    SetFill(sheet, row, numberColumn, MarketplaceFill);
                    marketplace.Add(key);
                    break;
                case IncomingKind.Collected:
                    SetFill(sheet, row, numberColumn, CollectedFill);
                    collected.Add(key);
                    break;
                case IncomingKind.NotCollected:
                    SetFill(sheet, row, numberColumn, NotCollectedFill);
                    notCollected.Add(key);
                    break;
                default:
                    SetFill(sheet, row, numberColumn, null);
                    unresolved.Add((number, cell, decision.Note));
                    break;
            }

            if (decision.Kind != IncomingKind.Unresolved && decision.Note.Length > 0)
            {
                warnings.Add(new ProcessingWarning(
                    "«" + number + "»: " + decision.Note + ".", name + ", " + cell, name, cell));
            }
        }

        _logger.Information(
            "«" + name + "»: удалено " + removed.Count + " (" +
            string.Join("; ", removed.GroupBy(d => d.Note.Split(' ')[0]).Select(g => g.Key + " - " + g.Count())) +
            "), маркетплейс " + marketplace.Count + ", собрано " + collected.Count +
            ", не собрано " + notCollected.Count + ", не разобрано " + unresolved.Count + ".");

        if (unresolved.Count > 0)
        {
            warnings.Add(new ProcessingWarning(
                "На листе «" + name + "» не получилось обработать строк: " + unresolved.Count + " (" +
                string.Join(", ", unresolved.Select(u => u.Number)) + "). Они остались без цвета " +
                "и не попали ни на один лист поставок - решите по ним вручную.",
                name));

            foreach (var (number, cell, note) in unresolved)
            {
                warnings.Add(new ProcessingWarning("«" + number + "»: " + note + ".", name + ", " + cell, name, cell));
            }
        }

        return new IncomingOutcome(collected, notCollected, marketplace, unresolved.Count);
    }

    private static IReadOnlySet<string> ReadPlanNumbers(object sheet)
    {
        var table = Headers(sheet, ReceivingSchema.Plan.Specs);
        var column = table.Headers[ReceivingSchema.Plan.Supplies];
        var first = table.Headers.HeaderRow + 1;
        var last = LastRow(sheet, table.Bounds, column);
        if (last < first)
        {
            return new HashSet<string>(StringComparer.Ordinal);
        }

        var grid = ExcelSheetOperations.ReadBlock(sheet, first, last, column, column, withFormulas: false);
        return SupplyNumbers.ExtractAll(Enumerable.Range(first, last - first + 1).Select(row => grid.Text(row, column)));
    }

    private static IEnumerable<string> AllTexts(object sheet)
    {
        var grid = ExcelSheetOperations.ReadGrid(sheet, withFormulas: false);
        for (var row = grid.FirstRow; row <= grid.LastRow; row++)
        {
            for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
            {
                yield return grid.Text(row, column);
            }
        }
    }

    private static NotAcceptedTable ReadNotAccepted(object sheet)
    {
        var table = Headers(sheet, ReceivingSchema.NotAccepted.Specs);
        var headerRow = table.Headers.HeaderRow;
        var headerGrid = ExcelSheetOperations.ReadBlock(sheet, headerRow, headerRow, 1, table.LastHeaderColumn, withFormulas: false);

        var columns = Enumerable.Range(1, table.LastHeaderColumn)
            .Select(column => (Name: headerGrid.NormalizedText(headerRow, column), Column: column))
            .Where(pair => pair.Name.Length > 0)
            .ToList();

        var numberColumn = table.Headers[ReceivingSchema.NotAccepted.SupplyNumber];
        var first = headerRow + 1;
        var last = LastRow(sheet, table.Bounds, numberColumn);
        var grid = last >= first
            ? ExcelSheetOperations.ReadBlock(sheet, first, last, 1, table.LastHeaderColumn, withFormulas: false)
            : null;

        return new NotAcceptedTable(
            columns,
            grid,
            first,
            last,
            numberColumn,
            table.Headers[ReceivingSchema.NotAccepted.Code],
            table.Headers[ReceivingSchema.NotAccepted.Deviation]);
    }

    /// <summary>
    /// Лист поставок: строки «Непринятого товара» по номерам поставок нужного цвета.
    /// Первые две колонки - формулы «Код» и «Тара/код поставщика» - протягиваются, остальные
    /// заполняются значениями по названиям колонок.
    /// </summary>
    private SuppliesOutcome WriteSupplies(
        object applicationObject,
        object sheet,
        NotAcceptedTable source,
        IReadOnlySet<string> keys,
        List<ProcessingWarning> warnings)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.Supplies.Specs);
        var headerRow = table.Headers.HeaderRow;
        var first = headerRow + 1;
        var headerGrid = ExcelSheetOperations.ReadBlock(sheet, headerRow, headerRow, 1, table.LastHeaderColumn, withFormulas: false);

        var mapping = new List<(int Source, int Target)>();
        var missing = new List<string>();
        foreach (var (sourceName, sourceColumn) in source.Columns)
        {
            var key = TextUtils.NormalizeKey(sourceName);
            var target = Enumerable.Range(ReceivingSchema.Supplies.FormulaColumns + 1, Math.Max(table.LastHeaderColumn - ReceivingSchema.Supplies.FormulaColumns, 0))
                .FirstOrDefault(column =>
                    string.Equals(TextUtils.NormalizeKey(headerGrid.Text(headerRow, column)), key, StringComparison.Ordinal) &&
                    mapping.All(pair => pair.Target != column));

            if (target > 0)
            {
                mapping.Add((sourceColumn, target));
            }
            else
            {
                missing.Add(sourceName);
            }
        }

        if (missing.Count > 0)
        {
            warnings.Add(new ProcessingWarning(
                "На листе «" + name + "» нет колонок " + string.Join(", ", missing.Select(m => "«" + m + "»")) +
                " - они с «Непринятого товара» не перенесены.",
                name));
        }

        var rows = new List<int>();
        if (source.Grid is not null)
        {
            for (var row = source.FirstRow; row <= source.LastRow; row++)
            {
                if (keys.Contains(SupplyNumbers.Key(source.Grid.NormalizedText(row, source.NumberColumn))))
                {
                    rows.Add(row);
                }
            }
        }

        var count = rows.Count;
        var width = new ColumnRange(1, table.LastHeaderColumn);
        var lastRow = Math.Max(first, first + count - 1);
        var oldLast = LastRow(sheet, table.Bounds, 1, table.Headers[ReceivingSchema.Supplies.SupplyNumber]);

        var template = ExcelSheetOperations.ReadBlock(
            sheet, first, first, 1, ReceivingSchema.Supplies.FormulaColumns, withFormulas: true);
        if (!Enumerable.Range(1, ReceivingSchema.Supplies.FormulaColumns).All(column => template.HasFormula(first, column)))
        {
            warnings.Add(new ProcessingWarning(
                "На листе «" + name + "» в первой строке нет формул «" + ReceivingSchema.Supplies.Code +
                "» и «" + ReceivingSchema.Supplies.ContainerCode + "»: протягивать нечего, и «итог» " +
                "не найдёт по этому листу тару и номер поставки. Верните формулы в первую строку.",
                name + ", " + new CellRef(first, 1),
                name,
                new CellRef(first, 1).ToString()));
        }

        ExcelSheetOperations.ClearRange(sheet, first + 1, oldLast, width.First, width.Last);
        FillDown(applicationObject, sheet, first, first + count - 1, width);

        foreach (var (sourceColumn, target) in mapping)
        {
            if (count == 0)
            {
                ExcelSheetOperations.ClearBlock(sheet, first, first, target);
                continue;
            }

            WriteColumn(sheet, first, target, rows.Select(row => source.Grid!.Value(row, sourceColumn)).ToList());
        }

        ReapplyAutoFilter(sheet, lastRow);

        var found = rows
            .Select(row => SupplyNumbers.Key(source.Grid!.NormalizedText(row, source.NumberColumn)))
            .ToHashSet(StringComparer.Ordinal);
        var absent = keys.Where(key => !found.Contains(key)).OrderBy(key => key, StringComparer.Ordinal).ToList();
        if (absent.Count > 0)
        {
            warnings.Add(new ProcessingWarning(
                "Поставки " + string.Join(", ", absent.Select(key => key.ToUpperInvariant())) + " покрашены в «Приходах», " +
                "но строк по ним на листе «" + ReceivingSchema.NotAcceptedSheet + "» нет - на «" + name + "» они не попали.",
                name));
        }

        _logger.Information("«" + name + "»: поставок " + found.Count + ", строк " + count + ".");

        var deviations = ReceivingStockRules.SumByCode(rows.Select(row =>
            (source.Grid!.Value(row, source.CodeColumn), source.Grid!.Value(row, source.DeviationColumn))));

        return new SuppliesOutcome(found.Count, deviations);
    }

    private static Dictionary<string, double> ReadAnswers(object sheet)
    {
        var table = Headers(sheet, ReceivingSchema.Agreement.Specs);
        var codeColumn = table.Headers[ReceivingSchema.Agreement.Code];
        var answerColumn = table.Headers[ReceivingSchema.Agreement.AnswerMarketplace];
        var first = table.Headers.HeaderRow + 1;
        var last = LastRow(sheet, table.Bounds, codeColumn);
        if (last < first)
        {
            return new Dictionary<string, double>(StringComparer.Ordinal);
        }

        var grid = ExcelSheetOperations.ReadBlock(sheet, first, last, 1, Math.Max(codeColumn, answerColumn), withFormulas: false);
        return ReceivingSummary.MarketplaceAnswers(
            Enumerable.Range(first, last - first + 1).Select(row => (grid.Value(row, codeColumn), grid.Value(row, answerColumn))));
    }

    /// <summary>
    /// «итог»: номер недели, строки «Остатков», по которым есть что-то в хранении или
    /// в поставках, формулы первой строки на все строки и «Количество МП» из «Согласования».
    /// Колонки, которые заполняет человек и второй этап, очищаются.
    /// </summary>
    private SummaryOutcome WriteSummary(
        object applicationObject,
        object sheet,
        StockRows stock,
        SummarySources sources,
        IReadOnlyDictionary<string, double> answers,
        int week,
        List<ProcessingWarning> warnings)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.Summary.Specs);
        var headers = table.Headers;
        var first = headers.HeaderRow + 1;
        var width = new ColumnRange(1, table.LastHeaderColumn);
        var codeColumn = headers[ReceivingSchema.Summary.Code];
        var oldLast = LastRow(sheet, table.Bounds, codeColumn, headers[ReceivingSchema.Summary.StorageQuantity]);

        var weekColumn = headers[ReceivingSchema.Summary.SeasonSharePlan];
        var weekCell = new CellRef(1, weekColumn).ToString();
        if (ExcelSheetOperations.ReadBlock(sheet, 1, 1, weekColumn, weekColumn, withFormulas: true).HasFormula(1, weekColumn))
        {
            warnings.Add(new ProcessingWarning(
                "На «" + name + "» в " + weekCell + " формула, а не номер недели: неделя не записана.",
                name + ", " + weekCell,
                name,
                weekCell));
        }
        else
        {
            ExcelSheetOperations.SetValue(sheet, 1, weekColumn, (double)week);
            _logger.Information("«" + name + "»: текущая неделя " + week + " записана в " + weekCell + ".");
        }

        var selected = ReceivingSummary.SelectRows(stock.Codes, sources);
        var count = selected.Count;
        var lastRow = Math.Max(first, first + count - 1);
        var template = ExcelSheetOperations.ReadBlock(sheet, first, first, 1, width.Last, withFormulas: true);

        ExcelSheetOperations.ClearRange(sheet, first + 1, oldLast, width.First, width.Last);
        FillDown(applicationObject, sheet, first, first + count - 1, width);

        var valueColumns = new HashSet<int>();
        for (var k = 0; k < ReceivingSchema.Summary.FromStock.Count; k++)
        {
            var column = headers[ReceivingSchema.Summary.FromStock[k]];
            valueColumns.Add(column);
            if (count == 0)
            {
                ExcelSheetOperations.ClearBlock(sheet, first, first, column);
                continue;
            }

            var index = k;
            WriteColumn(sheet, first, column, selected.Select(row => stock.Rows[row][index]).ToList());
        }

        foreach (var column in Enumerable.Range(width.First, width.Last - width.First + 1)
                     .Where(column => !valueColumns.Contains(column) && !template.HasFormula(first, column)))
        {
            ExcelSheetOperations.ClearBlock(sheet, first, lastRow, column);
        }

        var answered = 0;
        var answeredCodes = new HashSet<string>(StringComparer.Ordinal);
        var marketplaceColumn = headers[ReceivingSchema.Summary.QuantityMarketplace];
        for (var i = 0; i < count; i++)
        {
            var key = ReceivingStockRules.CodeKey(stock.Codes[selected[i]]);
            if (answers.TryGetValue(key, out var answer))
            {
                ExcelSheetOperations.SetValue(sheet, first + i, marketplaceColumn, answer);
                answeredCodes.Add(key);
                answered++;
            }
        }

        var lost = answers.Keys.Where(key => !answeredCodes.Contains(key)).ToList();
        if (lost.Count > 0)
        {
            warnings.Add(new ProcessingWarning(
                "В «" + ReceivingSchema.AgreementSheet + "» есть «Ответ МП» больше нуля по кодам, которых нет " +
                "на «итоге»: " + string.Join(", ", lost.Take(10)) + (lost.Count > 10 ? " и другие" : string.Empty) +
                ". По ним нет ни хранения, ни поставок - «Количество МП» записать некуда.",
                ReceivingSchema.AgreementSheet));
        }

        if (count == 0)
        {
            warnings.Add(new ProcessingWarning(
                "На «" + name + "» не осталось ни одной строки: ни по одному коду «Остатков» нет ни хранения, ни поставок.",
                name));
        }

        ReapplyAutoFilter(sheet, lastRow);
        _logger.Information(
            "«" + name + "»: строк " + count + " из " + stock.Rows.Count + " «Остатков», «Количество МП» из «Согласования» - " +
            answered + ".");

        return new SummaryOutcome(count, answered);
    }

    // ================= Этап 2: адреса =================

    private ProcessingResult PlaceAddresses(
        object applicationObject,
        object workbook,
        string path,
        IProgress<ProcessingStage>? progress,
        CancellationToken cancellationToken,
        ComScope scope)
    {
        dynamic application = applicationObject;

        var problems = new List<string>();
        var summarySheet = Required(workbook, ReceivingSchema.SummarySheet, scope, problems);
        var storageSheet = Required(workbook, ReceivingSchema.StorageSheet, scope, problems);
        var marketplaceSheet = Required(workbook, ReceivingSchema.MarketplaceSheet, scope, problems);
        var collectedSheet = Required(workbook, ReceivingSchema.CollectedSheet, scope, problems);
        var addressesSheet = Required(workbook, ReceivingSchema.AddressesSheet, scope, problems);
        var containersSheet = Required(workbook, ReceivingSchema.ContainersSheet, scope, problems);
        var marketplaceAddressesSheet = Required(workbook, ReceivingSchema.MarketplaceAddressesSheet, scope, problems);
        var marketplaceContainersSheet = Required(workbook, ReceivingSchema.MarketplaceContainersSheet, scope, problems);

        if (problems.Count > 0)
        {
            throw new WorkbookValidationException(problems);
        }

        // «Допоставить» заполняют с фильтром по сектору, и он так и остаётся в книге.
        // Под фильтром Excel не видит скрытых строк при поиске последней строки
        // и копирует только в видимые - поэтому сначала фильтры снимаются везде.
        Report(progress, "Снятие фильтров на всех листах", 6);
        ShowAllRowsEverywhere(workbook, scope);

        var warnings = new List<ProcessingWarning>();
        var summaryColumns = ReadSummaryColumns(summarySheet!);
        foreach (var sheet in new[] { storageSheet!, marketplaceSheet! })
        {
            RestoreToReceiveFormula(applicationObject, sheet, summaryColumns, warnings);
        }

        Report(progress, "Пересчёт формул", 10);
        application.CalculateFull();

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Листы «адреса» и «тары»", 25);
        var storageCounted = ReadCounted(storageSheet!);
        var addresses = AddressBlocks.Addresses(storageCounted);
        var containers = AddressBlocks.Containers(storageCounted);
        var addressLookup = WriteAddresses(applicationObject, addressesSheet!, addresses);
        var containerLookup = WriteContainers(applicationObject, containersSheet!, containers);

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "«итог»: место хранения и тара", 45);
        var collected = ReadCollectedLookup(collectedSheet!);
        var fill = FillSummary(summarySheet!, addressLookup, containerLookup, collected, warnings);

        Report(progress, "Пересчёт формул", 60);
        application.CalculateFull();

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, "Листы «адреса МП» и «тары МП»", 72);
        var marketplaceCounted = ReadCounted(marketplaceSheet!);
        var marketplaceAddresses = AddressBlocks.Addresses(marketplaceCounted);
        var marketplaceContainers = AddressBlocks.Containers(marketplaceCounted);
        WriteAddresses(applicationObject, marketplaceAddressesSheet!, marketplaceAddresses);
        WriteContainers(applicationObject, marketplaceContainersSheet!, marketplaceContainers);

        Report(progress, "Пересчёт формул", 88);
        application.Calculation = ExcelConstants.XlCalculationAutomatic;
        application.CalculateFull();

        return new ProcessingResult(
            path,
            new[]
            {
                new ProcessingCounter("адресов «А2, А3»", addresses.Count),
                new ProcessingCounter("тар «А2, А3»", containers.Count),
                new ProcessingCounter("место хранения с адресов", fill.FromAddresses),
                new ProcessingCounter("место хранения из поставок", fill.FromCollected),
                new ProcessingCounter("«Допоставить» из «Количество МП»", fill.FromMarketplace),
                new ProcessingCounter("адресов «МП»", marketplaceAddresses.Count),
                new ProcessingCounter("тар «МП»", marketplaceContainers.Count),
            },
            warnings);
    }

    /// <summary>
    /// «В приемку» на этапе адресов. Пустой она бывает в файлах, подготовленных версией,
    /// которая формулу не вписывала. Повторять ради этого подготовку нельзя - она заново
    /// собрала бы «итог» и стёрла решённое «Допоставить», - поэтому формула вписывается
    /// и протягивается здесь, до пересчёта.
    /// </summary>
    private void RestoreToReceiveFormula(
        object applicationObject, object sheet, SummaryColumns summary, List<ProcessingWarning> warnings)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.Storage.Specs);
        var headers = table.Headers;
        var first = headers.HeaderRow + 1;
        var column = headers[ReceivingSchema.Storage.ToReceive];
        var template = ExcelSheetOperations.ReadBlock(sheet, first, first, column, column, withFormulas: true);
        if (template.HasFormula(first, column))
        {
            return;
        }

        var codeColumn = headers[ReceivingSchema.Storage.Code];
        var last = LastRow(sheet, table.Bounds, headers[ReceivingSchema.Storage.Address], codeColumn);

        ExcelSheetOperations.SetFormula(
            sheet,
            first,
            column,
            ReceivingFormulas.ToReceive(summary.Sheet, summary.CodeColumn, summary.RestockColumn, codeColumn, first));
        FillDown(applicationObject, sheet, first, last, new ColumnRange(column, column));

        warnings.Add(ToReceiveRestored(
            name, first, column, TextUtils.Normalize(template.Text(first, column)), summary, codeColumn));
        _logger.Information("«" + name + "»: формула «В приемку» вписана и протянута до строки " + Math.Max(first, last) + ".");
    }

    private static ProcessingWarning ToReceiveRestored(
        string name, int row, int column, string found, SummaryColumns summary, int codeColumn)
    {
        var cell = new CellRef(row, column).ToString();
        return new ProcessingWarning(
            "На листе «" + name + "» в колонке «" + ReceivingSchema.Storage.ToReceive + "» не было формулы" +
            (found.Length > 0 ? " - в первой строке стояло «" + found + "»" : string.Empty) +
            ". Вписана " +
            ReceivingFormulas.ToReceiveLocal(summary.Sheet, summary.CodeColumn, summary.RestockColumn, codeColumn, row) +
            " и протянута на все строки. Скорее всего, при прошлой вставке со «Склада» данные съехали на одну колонку.",
            name + ", " + cell,
            name,
            cell);
    }

    private static SummaryColumns ReadSummaryColumns(object sheet)
    {
        var table = Headers(sheet, ReceivingSchema.Summary.Specs);
        return new SummaryColumns(
            ExcelSheetOperations.GetSheetName(sheet),
            table.Headers[ReceivingSchema.Summary.Code],
            table.Headers[ReceivingSchema.Summary.Restock]);
    }

    private static IReadOnlyList<CountedRow> ReadCounted(object sheet)
    {
        var table = Headers(sheet, ReceivingSchema.Storage.Specs);
        var headers = table.Headers;
        var first = headers.HeaderRow + 1;
        var counted = headers[ReceivingSchema.Storage.Counted];
        var last = LastRow(sheet, table.Bounds, headers[ReceivingSchema.Storage.Address], headers[ReceivingSchema.Storage.Code]);
        if (last < first)
        {
            return Array.Empty<CountedRow>();
        }

        var grid = ExcelSheetOperations.ReadBlock(sheet, first, last, 1, table.LastHeaderColumn, withFormulas: false);
        var rows = new List<CountedRow>();
        for (var row = first; row <= last; row++)
        {
            var value = grid.Value(row, counted);
            if (CellError.IsError(value) || !(TextUtils.CellToDouble(value) > 0))
            {
                continue;
            }

            rows.Add(new CountedRow(
                grid.Value(row, headers[ReceivingSchema.Storage.Address]),
                grid.Value(row, headers[ReceivingSchema.Storage.Container]),
                grid.Value(row, headers[ReceivingSchema.Storage.Code]),
                grid.Value(row, headers[ReceivingSchema.Storage.AddressNumber]),
                grid.Value(row, headers[ReceivingSchema.Storage.ContainerNumber])));
        }

        return rows;
    }

    private LookupColumns WriteAddresses(object applicationObject, object sheet, IReadOnlyList<AddressLine> lines)
    {
        var table = Headers(sheet, ReceivingSchema.Addresses.Specs);
        var headers = table.Headers;
        var columns = new[]
        {
            headers[ReceivingSchema.Addresses.Address],
            headers[ReceivingSchema.Addresses.Code],
            headers[ReceivingSchema.Addresses.AddressNumber],
            headers[ReceivingSchema.Addresses.Joined],
        };

        WriteLines(
            applicationObject,
            sheet,
            table,
            columns,
            lines.Count,
            new Func<int, object?>[]
            {
                i => lines[i].Address,
                i => lines[i].Code,
                i => lines[i].AddressNumber,
                i => lines[i].Joined.Length > 0 ? lines[i].Joined : null,
            });

        return new LookupColumns(
            ExcelSheetOperations.GetSheetName(sheet),
            headers[ReceivingSchema.Addresses.Code],
            headers[ReceivingSchema.Addresses.Joined]);
    }

    private LookupColumns WriteContainers(object applicationObject, object sheet, IReadOnlyList<ContainerLine> lines)
    {
        var table = Headers(sheet, ReceivingSchema.Containers.Specs);
        var headers = table.Headers;
        var columns = new[]
        {
            headers[ReceivingSchema.Containers.Address],
            headers[ReceivingSchema.Containers.Container],
            headers[ReceivingSchema.Containers.Code],
            headers[ReceivingSchema.Containers.AddressNumber],
            headers[ReceivingSchema.Containers.ContainerNumber],
            headers[ReceivingSchema.Containers.Joined],
        };

        WriteLines(
            applicationObject,
            sheet,
            table,
            columns,
            lines.Count,
            new Func<int, object?>[]
            {
                i => lines[i].Address,
                i => lines[i].Container,
                i => lines[i].Code,
                i => lines[i].AddressNumber,
                i => lines[i].ContainerNumber,
                i => lines[i].Joined.Length > 0 ? lines[i].Joined : null,
            });

        return new LookupColumns(
            ExcelSheetOperations.GetSheetName(sheet),
            headers[ReceivingSchema.Containers.Code],
            headers[ReceivingSchema.Containers.Joined]);
    }

    /// <summary>
    /// Общая запись листов «адреса» и «тары»: старые строки очищаются, формулы первой строки
    /// протягиваются, значения пишутся поверх. Последняя из колонок - перечень - пишется
    /// текстом: иначе «18189, 18166» Excel прочитал бы как число.
    /// </summary>
    private void WriteLines(
        object applicationObject,
        object sheet,
        SheetTable table,
        IReadOnlyList<int> columns,
        int count,
        IReadOnlyList<Func<int, object?>> values)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var first = table.Headers.HeaderRow + 1;
        var width = new ColumnRange(1, table.LastHeaderColumn);
        var oldLast = LastRow(sheet, table.Bounds, columns.ToArray());
        var lastRow = Math.Max(first, first + count - 1);

        ExcelSheetOperations.ClearRange(sheet, first + 1, oldLast, width.First, width.Last);
        FillDown(applicationObject, sheet, first, first + count - 1, width);

        var joined = columns[^1];
        ExcelSheetOperations.SetRangeAsText(sheet, first, lastRow, joined, joined);

        for (var k = 0; k < columns.Count; k++)
        {
            if (count == 0)
            {
                ExcelSheetOperations.ClearBlock(sheet, first, first, columns[k]);
                continue;
            }

            var select = values[k];
            var column = Enumerable.Range(0, count).Select(select).ToList();
            if (k == columns.Count - 1)
            {
                ExcelSheetOperations.SetColumnValues(sheet, first, columns[k], column);
            }
            else
            {
                WriteColumn(sheet, first, columns[k], column);
            }
        }

        ReapplyAutoFilter(sheet, lastRow);
        _logger.Information("«" + name + "»: записано строк " + count + ".");
    }

    private static CollectedLookup ReadCollectedLookup(object sheet)
    {
        var table = Headers(sheet, ReceivingSchema.Supplies.Specs);
        return new CollectedLookup(
            ExcelSheetOperations.GetSheetName(sheet),
            table.Headers[ReceivingSchema.Supplies.Code],
            table.Headers[ReceivingSchema.Supplies.SupplyNumber],
            table.Headers[ReceivingSchema.Supplies.ContainerCode],
            table.Headers[ReceivingSchema.Supplies.Barcode]);
    }

    /// <summary>
    /// «итог» второго этапа: место хранения и тара формулами ВПР - с «адресов» и «тар»
    /// или из «Поставок собраны», - и «Количество МП» в пустое «Допоставить».
    /// </summary>
    private SummaryFillOutcome FillSummary(
        object sheet,
        LookupColumns addresses,
        LookupColumns containers,
        CollectedLookup collected,
        List<ProcessingWarning> warnings)
    {
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var table = Headers(sheet, ReceivingSchema.Summary.Specs);
        var headers = table.Headers;
        var first = headers.HeaderRow + 1;
        var codeColumn = headers[ReceivingSchema.Summary.Code];
        var last = LastRow(sheet, table.Bounds, codeColumn);
        if (last < first)
        {
            warnings.Add(new ProcessingWarning("На «" + name + "» нет строк: сначала выполните подготовку.", name));
            return new SummaryFillOutcome(0, 0, 0);
        }

        var restockColumn = headers[ReceivingSchema.Summary.Restock];
        var placeColumn = headers[ReceivingSchema.Summary.Place];
        var containerColumn = headers[ReceivingSchema.Summary.ContainerCode];
        var barcodeColumn = headers[ReceivingSchema.Summary.Barcode];

        var grid = ExcelSheetOperations.ReadBlock(sheet, first, last, 1, table.LastHeaderColumn, withFormulas: false);
        var states = Enumerable.Range(0, last - first + 1)
            .Select(i => new SummaryState(
                i,
                grid.Value(first + i, restockColumn),
                grid.Value(first + i, headers[ReceivingSchema.Summary.Quantity]),
                grid.Value(first + i, headers[ReceivingSchema.Summary.QuantityMarketplace]),
                grid.Value(first + i, headers[ReceivingSchema.Summary.Collected])))
            .ToList();

        if (states.All(state => TextUtils.Normalize(TextUtils.CellToString(state.Restock)).Length == 0))
        {
            warnings.Add(new ProcessingWarning(
                "На «" + name + "» «" + ReceivingSchema.Summary.Restock + "» не заполнено ни в одной строке: " +
                "место хранения проставлено только по «Количеству», из собранных поставок - нигде.",
                name));
        }

        var fills = ReceivingSummary.PlanFills(states);
        int fromAddresses = 0, fromCollected = 0, fromMarketplace = 0;

        foreach (var fill in fills)
        {
            var row = first + fill.Index;
            var key = ExcelColumn.ToLetters(codeColumn) + row.ToString(CultureInfo.InvariantCulture);

            switch (fill.Place)
            {
                case PlaceSource.StorageAddresses:
                    ExcelSheetOperations.SetFormula(sheet, row, placeColumn, Lookup(key, addresses.Sheet, addresses.KeyColumn, addresses.ValueColumn));
                    ExcelSheetOperations.SetFormula(sheet, row, containerColumn, Lookup(key, containers.Sheet, containers.KeyColumn, containers.ValueColumn));
                    fromAddresses++;
                    break;
                case PlaceSource.CollectedSupply:
                    ExcelSheetOperations.SetFormula(sheet, row, placeColumn, Lookup(key, collected.Sheet, collected.KeyColumn, collected.NumberColumn));
                    ExcelSheetOperations.SetFormula(sheet, row, containerColumn, Lookup(key, collected.Sheet, collected.KeyColumn, collected.ContainerColumn));
                    ExcelSheetOperations.SetFormula(sheet, row, barcodeColumn, Lookup(key, collected.Sheet, collected.KeyColumn, collected.BarcodeColumn));
                    fromCollected++;
                    break;
            }

            if (fill.RestockFromMarketplace is { } restock)
            {
                ExcelSheetOperations.SetValue(sheet, row, restockColumn, restock);
                fromMarketplace++;
            }
        }

        _logger.Information(
            "«" + name + "»: место хранения с адресов " + fromAddresses + ", из собранных поставок " + fromCollected +
            ", «Допоставить» из «Количество МП» " + fromMarketplace + ".");

        return new SummaryFillOutcome(fromAddresses, fromCollected, fromMarketplace);
    }

    /// <summary>ВПР по коду: ключ - первая колонка диапазона, значение - последняя.</summary>
    private static string Lookup(string key, string sheet, int keyColumn, int valueColumn) =>
        "=VLOOKUP(" + key + ",'" + sheet.Replace("'", "''", StringComparison.Ordinal) + "'!" +
        ExcelColumn.ToLetters(keyColumn) + ":" + ExcelColumn.ToLetters(valueColumn) + "," +
        (valueColumn - keyColumn + 1).ToString(CultureInfo.InvariantCulture) + ",0)";

    // ================= Общее =================

    private sealed record SheetTable(HeaderMap Headers, SheetBounds Bounds, int LastHeaderColumn);

    private sealed record StockRows(IReadOnlyList<object?[]> Rows, IReadOnlyList<object?> Codes);

    private sealed record IncomingOutcome(
        IReadOnlySet<string> Collected,
        IReadOnlySet<string> NotCollected,
        IReadOnlySet<string> Marketplace,
        int Unresolved);

    private sealed record NotAcceptedTable(
        IReadOnlyList<(string Name, int Column)> Columns,
        SheetGrid? Grid,
        int FirstRow,
        int LastRow,
        int NumberColumn,
        int CodeColumn,
        int DeviationColumn);

    private sealed record SuppliesOutcome(int Supplies, IReadOnlyDictionary<string, double> DeviationByCode);

    private sealed record SummaryOutcome(int Rows, int MarketplaceAnswers);

    private sealed record LookupColumns(string Sheet, int KeyColumn, int ValueColumn);

    /// <summary>Где на «итоге» код и «Допоставить» - на них смотрит формула «В приемку».</summary>
    private sealed record SummaryColumns(string Sheet, int CodeColumn, int RestockColumn);

    private sealed record CollectedLookup(
        string Sheet, int KeyColumn, int NumberColumn, int ContainerColumn, int BarcodeColumn);

    private sealed record SummaryFillOutcome(int FromAddresses, int FromCollected, int FromMarketplace);

    private static object? Required(object workbook, string name, ComScope scope, List<string> problems)
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

    private static object? Optional(object workbook, string name, ComScope scope, List<string> problems)
    {
        var lookup = ExcelSheetOperations.FindSheet(workbook, name, scope);
        if (lookup.Sheet is null && lookup.Candidates.Count > 1)
        {
            problems.Add("в книге несколько листов с названием, начинающимся с «" + name + "»: " +
                         string.Join(", ", lookup.Candidates) + " - оставьте один");
        }

        return lookup.Sheet;
    }

    /// <summary>
    /// Строка заголовков, границы листа и последняя колонка с заголовком. Таблица кончается
    /// на последнем заголовке: занятый диапазон Excel бывает шире из-за оформления,
    /// и очищать или копировать тысячи пустых колонок незачем.
    /// </summary>
    private static SheetTable Headers(object sheet, IReadOnlyList<ColumnSpec> specs)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(sheet);
        var lastScan = Math.Min(bounds.FirstRow + HeaderScanRows - 1, bounds.LastRow);
        var grid = ExcelSheetOperations.ReadBlock(
            sheet, bounds.FirstRow, lastScan, 1, Math.Max(bounds.LastColumn, 1), withFormulas: false);
        var headers = HeaderResolver.Resolve(grid, ExcelSheetOperations.GetSheetName(sheet), specs, HeaderScanRows);

        var lastHeader = headers.Columns.Values.Max();
        for (var column = grid.LastColumn; column > lastHeader; column--)
        {
            if (grid.NormalizedText(headers.HeaderRow, column).Length > 0)
            {
                lastHeader = column;
                break;
            }
        }

        return new SheetTable(headers, bounds, lastHeader);
    }

    private static int LastRow(object sheet, SheetBounds bounds, params int[] columns) =>
        columns.Select(column => ExcelSheetOperations.GetLastFilledRow(sheet, column, bounds.LastRow)).DefaultIfEmpty(0).Max();

    private static int? FindColumn(object sheet, int headerRow, int lastColumn, IReadOnlyList<string> aliases)
    {
        var grid = ExcelSheetOperations.ReadBlock(sheet, headerRow, headerRow, 1, lastColumn, withFormulas: false);
        for (var column = 1; column <= lastColumn; column++)
        {
            var key = TextUtils.NormalizeKey(grid.Text(headerRow, column));
            if (aliases.Any(alias => key.StartsWith(alias, StringComparison.Ordinal)))
            {
                return column;
            }
        }

        return null;
    }

    /// <summary>Колонки справа налево подряд идущими кусками: (первая колонка, сколько).</summary>
    private static IEnumerable<(int Start, int Count)> DescendingRuns(IReadOnlyList<int> columns)
    {
        var ordered = columns.Distinct().OrderByDescending(column => column).ToList();
        var index = 0;
        while (index < ordered.Count)
        {
            var end = ordered[index];
            var start = end;
            index++;
            while (index < ordered.Count && ordered[index] == start - 1)
            {
                start = ordered[index];
                index++;
            }

            yield return (start, end - start + 1);
        }
    }

    private static Dictionary<string, double> QuantitiesByCode(IEnumerable<WarehouseRow> rows) =>
        ReceivingStockRules.SumByCode(rows.Select(row => (row.Code, row.Quantity)));

    private static bool Blank(object? value) => TextUtils.Normalize(TextUtils.CellToString(value)).Length == 0;

    /// <summary>
    /// Колонка значений. Текст, который Excel принял бы за число или дату («336-002», «35-38»,
    /// «198 512 901»), пишется в текстовом формате - иначе значение на листе изменилось бы.
    /// </summary>
    private static void WriteColumn(object sheet, int firstRow, int column, IReadOnlyList<object?> values)
    {
        if (values.Count == 0)
        {
            return;
        }

        if (values.Any(value => value is string text && CellError.LooksNumericToExcel(text)))
        {
            ExcelSheetOperations.SetRangeAsText(sheet, firstRow, firstRow + values.Count - 1, column, column);
        }

        ExcelSheetOperations.SetColumnValues(sheet, firstRow, column, values);
    }

    /// <summary>
    /// Строка-образец копируется на строки ниже вместе с формулами и оформлением -
    /// это и есть «протянуть формулы вниз».
    /// </summary>
    private static void FillDown(object applicationObject, object sheetObject, int templateRow, int lastRow, ColumnRange columns)
    {
        if (lastRow <= templateRow)
        {
            return;
        }

        dynamic application = applicationObject;
        dynamic sheet = sheetObject;

        using (var scope = new ComScope())
        {
            dynamic source = scope.Track(sheet.Range[Reference(templateRow, columns.First, templateRow, columns.Last)]);
            dynamic target = scope.Track(sheet.Range[Reference(templateRow + 1, columns.First, lastRow, columns.Last)]);
            source.Copy(target);
        }

        application.CutCopyMode = false;
    }

    /// <summary>
    /// Автофильтр растягивается на новые строки: Excel сам его не расширяет, а по «итогу»
    /// аналитик дальше фильтруется по секторам и колонкам.
    /// </summary>
    private void ReapplyAutoFilter(object sheetObject, int lastRow)
    {
        dynamic sheet = sheetObject;
        try
        {
            bool mode = sheet.AutoFilterMode;
            if (!mode)
            {
                return;
            }

            using var scope = new ComScope();
            dynamic filter = scope.Track(sheet.AutoFilter);
            dynamic range = scope.Track(filter.Range);
            int row = range.Row;
            int column = range.Column;
            dynamic rangeColumns = scope.Track(range.Columns);
            int count = rangeColumns.Count;

            if (lastRow <= row)
            {
                return;
            }

            sheet.AutoFilterMode = false;
            dynamic target = scope.Track(sheet.Range[Reference(row, column, lastRow, column + count - 1)]);
            target.AutoFilter();
        }
        catch (COMException ex)
        {
            _logger.Warning("Не удалось растянуть автофильтр на листе «" + ExcelSheetOperations.GetSheetName(sheetObject) + "».", ex);
        }
        catch (RuntimeBinderException ex)
        {
            _logger.Warning("Не удалось растянуть автофильтр на листе «" + ExcelSheetOperations.GetSheetName(sheetObject) + "».", ex);
        }
    }

    private void ShowAllRowsEverywhere(object workbookObject, ComScope scope)
    {
        dynamic workbook = workbookObject;
        dynamic sheets = scope.Track(workbook.Worksheets);
        int count = sheets.Count;
        var cleared = 0;

        for (var index = 1; index <= count; index++)
        {
            object sheet = sheets[index];
            try
            {
                if (ExcelSheetOperations.ShowAllRows(sheet))
                {
                    cleared++;
                }
            }
            catch (COMException ex)
            {
                _logger.Warning("Не удалось снять фильтр на листе «" + ExcelSheetOperations.GetSheetName(sheet) + "».", ex);
            }
            finally
            {
                ComUtils.Release(sheet);
            }
        }

        _logger.Information("Фильтры сняты на листах: " + cleared + ".");
    }

    /// <summary>
    /// Заливка ячейки. Без цвета - снимается только своя заливка: цвет, поставленный
    /// человеком, сохраняется.
    /// </summary>
    private static void SetFill(object sheetObject, int row, int column, int? color)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        dynamic interior = scope.Track(cell.Interior);

        if (color is { } value)
        {
            interior.Color = value;
            return;
        }

        object? current = interior.Color;
        if (current is not null && Array.IndexOf(OwnFills, Convert.ToInt32(current, CultureInfo.InvariantCulture)) >= 0)
        {
            interior.ColorIndex = ExcelConstants.XlColorIndexNone;
        }
    }

    private static string Reference(int firstRow, int firstColumn, int lastRow, int lastColumn) =>
        ExcelColumn.ToLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture) + ":" +
        ExcelColumn.ToLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture);

    private static void Report(IProgress<ProcessingStage>? progress, string message, int percent) =>
        progress?.Report(new ProcessingStage(message, percent));
}
