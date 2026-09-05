using System.Globalization;
using System.Runtime.InteropServices;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Обработка книги через COM-автоматизацию установленного Microsoft Excel.
/// Исходный файл никогда не изменяется: все действия выполняются над копией.
/// </summary>
public sealed class ExcelWorkbookProcessor : IWorkbookProcessor
{
    /// <summary>Сколько первых строк листа просматривается в поиске строки заголовков.</summary>
    private const int HeaderScanRows = 15;

    /// <summary>Размер порции строк при чтении больших листов.</summary>
    private const int ChunkRows = 50000;

    private readonly IAppLogger _logger;
    private readonly Func<DateTime> _nowProvider;

    public ExcelWorkbookProcessor(IAppLogger logger, Func<DateTime>? nowProvider = null)
    {
        _logger = logger;
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

        var outputPath = BuildOutputPath(sourceFilePath);
        _logger.Information("Начало обработки. Исходный файл: " + sourceFilePath);
        Report(progress, "Создание копии файла", 5);

        File.Copy(sourceFilePath, outputPath);
        _logger.Information("Создана копия: " + outputPath);

        try
        {
            var result = ProcessCopy(outputPath, progress, cancellationToken);
            _logger.Information(
                "Обработка завершена. Обработано строк выгрузки: " + result.ProcessedOrderRows +
                ", осталось для ручной проверки: " + result.RemainingOrderRows +
                ", добавлено новых заказов: " + result.NewPlanOrders);
            return result;
        }
        catch (Exception ex)
        {
            _logger.Error("Обработка прервана, частичный результат удалён.", ex);
            TryDeleteFile(outputPath);
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

            var sheets = ResolveSheets((object)workbook, scope);
            cancellationToken.ThrowIfCancellationRequested();

            Report(progress, "Чтение данных", 15);
            var listSeparator = ExcelSheetOperations.GetListSeparator(applicationObject);

            // Отбор автофильтра снимается до чтения: при активном отборе поиск последней
            // заполненной строки останавливается на последней видимой, и часть данных
            // не попала бы в обработку. Удаление строк при активном отборе тоже неполное.
            if (ExcelSheetOperations.ShowAllRows(sheets.Orders))
            {
                _logger.Information("На листе «" + SheetSchema.OrdersSheet + "» снят отбор автофильтра.");
            }

            if (ExcelSheetOperations.ShowAllRows(sheets.Plan))
            {
                _logger.Information("На листе «" + SheetSchema.PlanSheet + "» снят отбор автофильтра.");
            }

            var orders = ReadOrdersSheet(sheets.Orders);
            var journal = ReadJournalSheet(sheets.Journal);
            var planLayout = ReadPlanLayout(sheets.Plan);
            _logger.Information(
                "Прочитано строк: выгрузка " + orders.Rows.Count +
                ", журнал " + journal.Rows.Count +
                ", план " + planLayout.AllDataRows.Count());

            var originalAggregates = planLayout.OrderSections.ToDictionary(
                section => section.Kind,
                section => (section.AggregateFormula, section.FirstDataRow));

            // Все перестроения листа «План» идут внутри колонок таблицы, поэтому высоты
            // строк за содержимым не следуют и восстанавливаются в конце по образцу.
            var planColumns = new ColumnRange(
                planLayout.Headers.Columns.Values.Min(),
                planLayout.Headers.Columns.Values.Max());
            var rowHeights = CaptureRowHeights(sheets.Plan, planLayout);

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Классификация заказов", 30);
            var classification = OrderClassifier.Classify(orders.Rows);
            var update = PlanUpdateBuilder.Build(planLayout, classification, journal.Rows, _nowProvider());
            _logger.Information(
                "Классификация: к удалению " + classification.RowsToDelete.Count +
                ", остаётся " + classification.Leftovers.Count +
                ", сегодняшних заказов " + classification.Groups.Count +
                ", новых строк плана " + update.NewRows.Count);

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Удаление обработанных строк выгрузки", 45);
            ExcelSheetOperations.DeleteRows(sheets.Orders, classification.RowsToDelete, listSeparator);

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Обновление листа «План»", 60);

            // Подстановка номеров выполняется до удаления и вставки строк: адреса строк
            // рассчитаны по той же разметке, которую видел PlanUpdateBuilder.
            ApplyPlannedMatches(sheets.Plan, planLayout, update.PlannedMatches);

            ExcelSheetOperations.DeleteRows(sheets.Plan, update.PlanRowsToDelete, listSeparator, planColumns);

            InsertNewRows(applicationObject, sheets.Plan, update.NewRows, planColumns);
            planLayout = ReadPlanLayout(sheets.Plan);

            ApplyOrderUpdates(sheets.Plan, planLayout, update.OrderUpdates);
            ApplyAggregateUpdates(sheets.Plan, planLayout, update.AggregateUpdates);

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Пересчёт формул", 75);
            application.Calculation = ExcelConstants.XlCalculationAutomatic;
            application.CalculateFull();

            Report(progress, "Нумерация строк", 85);
            planLayout = ReadPlanLayout(sheets.Plan);
            var numberColumn = planLayout.Headers[SheetSchema.Plan.Number];
            foreach (var assignment in PlanNumberingBuilder.Build(planLayout))
            {
                ExcelSheetOperations.SetValue(sheets.Plan, assignment.ExcelRow, numberColumn, (double)assignment.Number);
            }

            RepairAggregateFormulas(sheets.Plan, originalAggregates);
            ApplyRowHeights(sheets.Plan, ReadPlanLayout(sheets.Plan), rowHeights);

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Сохранение файла", 95);
            application.CalculateFull();
            workbook.Save();
            workbook.Close(true);
            closed = true;

            Report(progress, "Готово", 100);
            return new ProcessingResult(
                path,
                classification.RowsToDelete.Count,
                classification.Leftovers.Count,
                update.NewRows.Count,
                update.PlanRowsToDelete.Count);
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

    private sealed class WorkbookSheets
    {
        public WorkbookSheets(object plan, object orders, object journal)
        {
            Plan = plan;
            Orders = orders;
            Journal = journal;
        }

        public object Plan { get; }

        public object Orders { get; }

        public object Journal { get; }
    }

    private static WorkbookSheets ResolveSheets(object workbook, ComScope scope)
    {
        var plan = ExcelSheetOperations.FindSheet(workbook, SheetSchema.PlanSheet, scope);
        var orders = ExcelSheetOperations.FindSheet(workbook, SheetSchema.OrdersSheet, scope);
        var journal = ExcelSheetOperations.FindSheet(workbook, SheetSchema.JournalSheet, scope);

        var problems = new List<string>();
        Describe(plan, SheetSchema.PlanSheet, problems);
        Describe(orders, SheetSchema.OrdersSheet, problems);
        Describe(journal, SheetSchema.JournalSheet, problems);

        if (problems.Count > 0)
        {
            throw new WorkbookValidationException(problems);
        }

        return new WorkbookSheets(plan.Sheet!, orders.Sheet!, journal.Sheet!);
    }

    private static void Describe(SheetLookup lookup, string name, List<string> problems)
    {
        if (lookup.Sheet is not null)
        {
            return;
        }

        problems.Add(lookup.Candidates.Count == 0
            ? "в книге нет листа, название которого начинается с «" + name + "»"
            : "в книге несколько листов с названием, начинающимся с «" + name + "»: " +
              string.Join(", ", lookup.Candidates) + " - оставьте один");
    }

    private static PlanLayout ReadPlanLayout(object planSheet) =>
        PlanSheetReader.Read(ExcelSheetOperations.ReadGrid(planSheet, withFormulas: true));

    /// <summary>
    /// Читает выгрузку заказов по частям. Из листа берутся только колонки, которые нужны
    /// правилам, и только строки до последней заполненной: это не даёт большому файлу
    /// целиком оказаться в памяти и отсекает раздутый использованный диапазон.
    /// </summary>
    private OrdersSheet ReadOrdersSheet(object ordersSheet)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(ordersSheet);
        var headerGrid = ExcelSheetOperations.ReadBlock(
            ordersSheet,
            bounds.FirstRow,
            Math.Min(bounds.FirstRow + HeaderScanRows - 1, bounds.LastRow),
            bounds.FirstColumn,
            bounds.LastColumn,
            withFormulas: false);

        var headers = OrdersSheetReader.ResolveHeaders(headerGrid);
        var firstColumn = headers.Columns.Values.Min();
        var lastColumn = headers.Columns.Values.Max();

        var lastRow = Math.Min(
            bounds.LastRow,
            Math.Max(
                ExcelSheetOperations.GetLastFilledRow(
                    ordersSheet, headers[SheetSchema.Orders.Comment], bounds.LastRow),
                ExcelSheetOperations.GetLastFilledRow(
                    ordersSheet, headers[SheetSchema.Orders.Division], bounds.LastRow)));

        var rows = new List<OrderRow>();
        for (var start = headers.HeaderRow + 1; start <= lastRow; start += ChunkRows)
        {
            var end = Math.Min(start + ChunkRows - 1, lastRow);
            var chunk = ExcelSheetOperations.ReadBlock(
                ordersSheet, start, end, firstColumn, lastColumn, withFormulas: false);
            OrdersSheetReader.ReadRows(chunk, headers, rows);
        }

        return new OrdersSheet(headers, rows);
    }

    /// <summary>Читает журнал по тем же правилам. Порядок строк сохраняется.</summary>
    private JournalSheet ReadJournalSheet(object journalSheet)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(journalSheet);
        var headerGrid = ExcelSheetOperations.ReadBlock(
            journalSheet,
            bounds.FirstRow,
            Math.Min(bounds.FirstRow + HeaderScanRows - 1, bounds.LastRow),
            bounds.FirstColumn,
            bounds.LastColumn,
            withFormulas: false);

        var headers = JournalSheetReader.ResolveHeaders(headerGrid);
        var firstColumn = headers.Columns.Values.Min();
        var lastColumn = headers.Columns.Values.Max();

        var lastRow = Math.Min(
            bounds.LastRow,
            ExcelSheetOperations.GetLastFilledRow(
                journalSheet, headers[SheetSchema.Journal.Comment], bounds.LastRow));

        var rows = new List<JournalRow>();
        var order = 0;
        for (var start = headers.HeaderRow + 1; start <= lastRow; start += ChunkRows)
        {
            var end = Math.Min(start + ChunkRows - 1, lastRow);
            var chunk = ExcelSheetOperations.ReadBlock(
                journalSheet, start, end, firstColumn, lastColumn, withFormulas: false);
            JournalSheetReader.ReadRows(chunk, headers, rows, ref order);
        }

        return new JournalSheet(headers, rows);
    }

    /// <summary>
    /// Высоты строк листа «План» по их роли. Перестроения идут внутри колонок таблицы,
    /// поэтому высота остаётся за номером строки, а не за содержимым, и после всех
    /// перемещений её нужно расставить заново.
    /// </summary>
    private sealed record PlanRowHeights(double? Section, double? Special, double? Data, double? Marketplace);

    private static PlanRowHeights CaptureRowHeights(object planSheet, PlanLayout layout)
    {
        double? Height(int? row) => row is null ? null : ExcelSheetOperations.GetRowHeight(planSheet, row.Value);

        var firstSection = layout.Sections.Count > 0 ? layout.Sections[0].HeaderRow : (int?)null;
        var firstOrderRow = layout.OrderSections
            .SelectMany(s => s.DataRows)
            .FirstOrDefault(r => r.IsOrderRow)?.ExcelRow;
        var firstMarketplaceRow = layout.Section(PlanSectionKind.Marketplaces)?.DataRows.FirstOrDefault()?.ExcelRow;

        return new PlanRowHeights(
            Height(firstSection),
            Height(layout.StorageAcceptanceRow?.ExcelRow),
            Height(firstOrderRow),
            Height(firstMarketplaceRow));
    }

    private void ApplyRowHeights(object planSheet, PlanLayout layout, PlanRowHeights heights)
    {
        void Set(int row, double? height)
        {
            if (height is > 0d)
            {
                ExcelSheetOperations.SetRowHeight(planSheet, row, height.Value);
            }
        }

        foreach (var section in layout.Sections)
        {
            Set(section.HeaderRow, heights.Section);

            foreach (var row in section.DataRows)
            {
                if (section.Kind == PlanSectionKind.Marketplaces)
                {
                    Set(row.ExcelRow, heights.Marketplace);
                }
                else if (!row.IsOrderRow)
                {
                    Set(row.ExcelRow, heights.Special);
                }
                else
                {
                    Set(row.ExcelRow, heights.Data);
                }
            }
        }

        _logger.Debug("Высоты строк листа «План» восстановлены по образцу.");
    }

    private void InsertNewRows(
        object application,
        object planSheet,
        IReadOnlyList<NewPlanRowSpec> specs,
        ColumnRange columns)
    {
        foreach (var spec in specs)
        {
            var layout = ReadPlanLayout(planSheet);
            var section = layout.Section(spec.Section);
            if (section is null)
            {
                throw new WarehousePlanException(
                    "На листе «" + SheetSchema.PlanSheet + "» не найден блок для нового заказа " +
                    spec.LoadNumber + ".");
            }

            var template = ChooseTemplateRow(layout, section);
            if (template is null)
            {
                throw new WarehousePlanException(
                    "На листе «" + SheetSchema.PlanSheet + "» нет ни одной строки-образца, " +
                    "по которой можно построить новую строку заказа.");
            }

            // Новая строка встаёт под последнюю строку блока: порядок строк, который
            // выстроила аналитик, программа не трогает, а всё сегодняшнее собирается
            // внизу блока, где его видно одним куском.
            var insertBefore = section.LastDataRow + 1;
            ExcelSheetOperations.InsertCopiedRow(
                application, planSheet, template.ExcelRow, insertBefore, columns);
            FillNewRow(planSheet, layout.Headers, insertBefore, spec);
            _logger.Debug("Добавлена строка заказа " + spec.LoadNumber + " в строку " + insertBefore + ".");
        }
    }

    /// <summary>
    /// Строка-образец выбирается внутри того же блока. Предпочтение отдаётся нижней строке
    /// реального заказа, у которой «Дата в сети» задана формулой: тогда новая строка получает
    /// такие же формулы и оформление, как соседние строки блока.
    /// </summary>
    private static PlanRow? ChooseTemplateRow(PlanLayout layout, PlanSection section)
    {
        var candidates = section.DataRows.Where(row => row.IsOrderRow).ToList();
        if (candidates.Count == 0)
        {
            candidates = layout.OrderSections
                .SelectMany(s => s.DataRows)
                .Where(row => row.IsOrderRow)
                .ToList();
        }

        if (candidates.Count == 0)
        {
            return null;
        }

        bool HasNetworkDateFormula(PlanRow row) => row.FormulaColumns.Contains(SheetSchema.Plan.NetworkDate);

        return candidates.LastOrDefault(row => HasNetworkDateFormula(row) && row.LoadNumber.HasValue)
               ?? candidates.LastOrDefault(HasNetworkDateFormula)
               ?? candidates.LastOrDefault(row => row.LoadNumber.HasValue)
               ?? candidates[^1];
    }

    /// <summary>
    /// Заполняет новую строку. Скопированные из образца формулы сохраняются,
    /// значения, не заданные техническим заданием, очищаются.
    /// </summary>
    private static void FillNewRow(object planSheet, HeaderMap headers, int row, NewPlanRowSpec spec)
    {
        var firstColumn = headers.Columns.Values.Min();
        var lastColumn = headers.Columns.Values.Max();
        var formulas = ExcelSheetOperations.ReadRowFormulas(planSheet, row, firstColumn, lastColumn);

        bool HasFormula(int column)
        {
            var index = column - firstColumn;
            if (index < 0 || index >= formulas.Length)
            {
                return false;
            }

            var formula = formulas[index];
            return formula is not null && formula.StartsWith("=", StringComparison.Ordinal);
        }

        foreach (var pair in headers.Columns)
        {
            var column = pair.Value;
            if (HasFormula(column))
            {
                continue;
            }

            switch (pair.Key)
            {
                case SheetSchema.Plan.Supplies:
                    ExcelSheetOperations.SetValue(planSheet, row, column, spec.Supplies);
                    break;

                case SheetSchema.Plan.Processing:
                    if (spec.Processing.Length > 0)
                    {
                        ExcelSheetOperations.SetValue(planSheet, row, column, spec.Processing);
                    }
                    else
                    {
                        ExcelSheetOperations.ClearValue(planSheet, row, column);
                    }

                    break;

                case SheetSchema.Plan.Comments:
                    ExcelSheetOperations.SetValue(planSheet, row, column, spec.Comments);
                    break;

                case SheetSchema.Plan.DocumentDate:
                    if (spec.DocumentDate.HasValue)
                    {
                        ExcelSheetOperations.SetValue(planSheet, row, column, spec.DocumentDate.Value);
                    }
                    else
                    {
                        ExcelSheetOperations.ClearValue(planSheet, row, column);
                    }

                    break;

                case SheetSchema.Plan.LoadNumber:
                    ExcelSheetOperations.SetValue(planSheet, row, column, (double)spec.LoadNumber);
                    break;

                case SheetSchema.Plan.Quantity:
                    ExcelSheetOperations.SetValue(planSheet, row, column, 0d);
                    break;

                case SheetSchema.Plan.CompletionPercent:
                    ExcelSheetOperations.SetValue(planSheet, row, column, 0d);
                    break;

                default:
                    ExcelSheetOperations.ClearValue(planSheet, row, column);
                    break;
            }
        }
    }

    /// <summary>
    /// Вписывает номер загрузки в строку, заведённую заранее без него. Аналитик планирует
    /// будущую поставку отдельной строкой, и когда заказ приходит - дополняет её,
    /// а не заводит вторую такую же.
    ///
    /// Трогаются только те поля, которые меняет и она: номер загрузки и комментарий.
    /// «Обработка» проставляется, лишь если у поставки есть номер и ячейка пустая:
    /// значение, выставленное вручную, не затирается. Количество, статус и процент
    /// проставит обычное обновление заказов - строку оно найдёт уже по номеру.
    /// </summary>
    private void ApplyPlannedMatches(
        object planSheet,
        PlanLayout layout,
        IReadOnlyList<PlannedRowMatch> matches)
    {
        if (matches.Count == 0)
        {
            return;
        }

        var loadNumberColumn = layout.Headers[SheetSchema.Plan.LoadNumber];
        var commentsColumn = layout.Headers[SheetSchema.Plan.Comments];
        var processingColumn = layout.Headers[SheetSchema.Plan.Processing];
        var byRow = layout.AllDataRows.ToDictionary(row => row.ExcelRow);

        foreach (var match in matches)
        {
            ExcelSheetOperations.SetValue(planSheet, match.ExcelRow, loadNumberColumn, (double)match.LoadNumber);
            ExcelSheetOperations.SetValue(planSheet, match.ExcelRow, commentsColumn, OrderTextRules.LoadedComment);

            if (match.Processing.Length > 0 &&
                byRow.TryGetValue(match.ExcelRow, out var row) &&
                row.Processing.Length == 0)
            {
                ExcelSheetOperations.SetValue(planSheet, match.ExcelRow, processingColumn, match.Processing);
            }

            _logger.Information(
                "Заказ " + match.LoadNumber + " вписан в запланированную строку " + match.ExcelRow + ".");
        }
    }

    private static void ApplyOrderUpdates(
        object planSheet,
        PlanLayout layout,
        IReadOnlyList<OrderRowUpdate> updates)
    {
        if (updates.Count == 0)
        {
            return;
        }

        var byLoadNumber = new Dictionary<long, OrderRowUpdate>();
        foreach (var update in updates)
        {
            byLoadNumber[update.LoadNumber] = update;
        }

        var quantityColumn = layout.Headers[SheetSchema.Plan.Quantity];
        var statusColumn = layout.Headers[SheetSchema.Plan.Status];
        var percentColumn = layout.Headers[SheetSchema.Plan.CompletionPercent];
        var loadNumberColumn = layout.Headers[SheetSchema.Plan.LoadNumber];

        foreach (var row in layout.OrderSections.SelectMany(section => section.DataRows))
        {
            if (!row.IsOrderRow || row.LoadNumber is null)
            {
                continue;
            }

            if (!byLoadNumber.TryGetValue(row.LoadNumber.Value, out var update))
            {
                continue;
            }

            // Номер загрузки подсвечивается: зелёным - строка, добавленная сегодня,
            // розовым - заказ, который пропал из выгрузки и которого в плане быть уже
            // не должно. Строка прошлого дня свою пометку теряет.
            var mark = update.MissingFromOrders
                ? RowMark.Missing
                : update.IsNewRow ? RowMark.Added : RowMark.None;
            ExcelSheetOperations.SetRowMark(planSheet, row.ExcelRow, loadNumberColumn, mark);

            // Формулы не заменяются значениями: если пользователь считает какое-то из
            // этих полей формулой, она сохраняется.
            if (!row.FormulaColumns.Contains(SheetSchema.Plan.Quantity))
            {
                ExcelSheetOperations.SetValue(planSheet, row.ExcelRow, quantityColumn, update.Quantity);
            }

            if (update.Status is not null &&
                !row.FormulaColumns.Contains(SheetSchema.Plan.Status) &&
                !IsAlreadyInAssembly(row.Status))
            {
                ExcelSheetOperations.SetValue(planSheet, row.ExcelRow, statusColumn, update.Status);
            }

            if (update.CompletionPercent is not null &&
                !row.FormulaColumns.Contains(SheetSchema.Plan.CompletionPercent))
            {
                ExcelSheetOperations.SetValue(planSheet, row.ExcelRow, percentColumn, update.CompletionPercent.Value);
            }
        }
    }

    /// <summary>
    /// В книге статус встречается и как «в сборке», и как «в сборке.» с точкой.
    /// Если строка уже помечена как «в сборке», её текст не переписывается:
    /// смысл не меняется, а привычное пользователю написание сохраняется.
    /// </summary>
    private static bool IsAlreadyInAssembly(string? status) =>
        TextUtils.NormalizeKey(status).TrimEnd('.') == OrderTextRules.InAssemblyStatus;

    private void ApplyAggregateUpdates(
        object planSheet,
        PlanLayout layout,
        IReadOnlyList<AggregateUpdate> updates)
    {
        var quantityColumn = layout.Headers[SheetSchema.Plan.Quantity];

        foreach (var update in updates)
        {
            var targetRow = ResolveAggregateRow(layout, update.Target);
            if (targetRow is null)
            {
                _logger.Warning("Не найдена строка для итога " + update.Target + ", значение не записано.");
                continue;
            }

            ExcelSheetOperations.SetValue(planSheet, targetRow.Value, quantityColumn, update.Quantity);
        }
    }

    private static int? ResolveAggregateRow(PlanLayout layout, PlanAggregateTarget target) => target switch
    {
        PlanAggregateTarget.AutoHub => layout.AutoHubRow?.ExcelRow,
        PlanAggregateTarget.Wholesale => layout.Section(PlanSectionKind.Wholesale)?.HeaderRow,
        PlanAggregateTarget.InternetShop => layout.Section(PlanSectionKind.InternetShop)?.HeaderRow,
        PlanAggregateTarget.MarketplaceFromStorage => layout.MarketplaceFromStorage?.ExcelRow,
        PlanAggregateTarget.MarketplaceFromReturns => layout.MarketplaceFromReturns?.ExcelRow,
        PlanAggregateTarget.MarketplaceFromSupplies => layout.MarketplaceFromSupplies?.ExcelRow,
        _ => null,
    };

    private void RepairAggregateFormulas(
        object planSheet,
        IReadOnlyDictionary<PlanSectionKind, (string? AggregateFormula, int FirstDataRow)> original)
    {
        var layout = ReadPlanLayout(planSheet);
        var quantityColumn = layout.Headers[SheetSchema.Plan.Quantity];

        foreach (var section in layout.OrderSections)
        {
            if (section.DataRows.Count == 0 ||
                !original.TryGetValue(section.Kind, out var info) ||
                info.AggregateFormula is null)
            {
                continue;
            }

            var repaired = AggregateFormulaRepair.BuildRepairedFormula(
                info.AggregateFormula,
                info.FirstDataRow,
                section.FirstDataRow,
                section.LastDataRow);

            if (repaired is null)
            {
                continue;
            }

            var current = ExcelSheetOperations.GetFormula(planSheet, section.HeaderRow, quantityColumn);
            if (string.Equals(current, repaired, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            ExcelSheetOperations.SetFormula(planSheet, section.HeaderRow, quantityColumn, repaired);
            _logger.Debug("Итоговая формула блока " + section.Kind + " приведена к виду " + repaired + ".");
        }
    }

    private string BuildOutputPath(string sourceFilePath)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(sourceFilePath));
        if (string.IsNullOrEmpty(directory))
        {
            throw new WarehousePlanException("Не удалось определить папку исходного файла.");
        }

        var name = Path.GetFileNameWithoutExtension(sourceFilePath);
        var extension = Path.GetExtension(sourceFilePath);
        var stamp = _nowProvider().ToString("yyyy-MM-dd_HHmmss", CultureInfo.InvariantCulture);

        var candidate = Path.Combine(directory, name + "_готово_" + stamp + extension);
        var attempt = 1;
        while (File.Exists(candidate))
        {
            candidate = Path.Combine(
                directory,
                name + "_готово_" + stamp + "_" + attempt.ToString(CultureInfo.InvariantCulture) + extension);
            attempt++;
        }

        return candidate;
    }

    private void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException ex)
        {
            _logger.Warning("Не удалось удалить частичный результат " + path + ".", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            _logger.Warning("Нет прав на удаление частичного результата " + path + ".", ex);
        }
    }

    private static void Report(IProgress<ProcessingStage>? progress, string message, int percent) =>
        progress?.Report(new ProcessingStage(message, percent));
}
