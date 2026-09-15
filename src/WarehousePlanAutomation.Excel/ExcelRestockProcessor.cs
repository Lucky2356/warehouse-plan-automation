using System.Runtime.InteropServices;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Подтоварка маркетплейса: колонка «Запрет» на листе «на загрузку», а за ней - места
/// хранения и загрузочники (<see cref="ExcelRestockStorageStage"/>).
///
/// Программа не ходит по сети: «Запрет», «Фактическое кол-во» и «Прогнозный sellout»
/// аналитик подтягивает из сезонного файла сама, а программа их читает, принимает
/// по ним решение и складывает всё спорное на отдельный лист.
///
/// Исходный файл никогда не изменяется: все действия выполняются над копией.
/// </summary>
public sealed class ExcelRestockProcessor : IWorkbookProcessor
{
    private readonly IAppLogger _logger;
    private readonly Func<DateTime> _nowProvider;

    public ExcelRestockProcessor(IAppLogger logger, Func<DateTime>? nowProvider = null)
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

        var outputPath = OutputPath.Build(sourceFilePath, _nowProvider(), OutputFileName.RestockMark);
        _logger.Information("Начало разбора подтоварки. Исходный файл: " + sourceFilePath);
        Report(progress, "Создание копии файла", 5);

        File.Copy(sourceFilePath, outputPath);
        _logger.Information("Создана копия: " + outputPath);

        try
        {
            var result = ProcessCopy(outputPath, progress, cancellationToken);
            _logger.Information(
                "Подтоварка разобрана. " +
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

            var lookup = ExcelSheetOperations.FindSheet((object)workbook, RestockSchema.LoadSheet, scope);
            if (lookup.Sheet is null)
            {
                throw new WorkbookValidationException(new[]
                {
                    lookup.Candidates.Count == 0
                        ? "в книге нет листа, название которого начинается с «" + RestockSchema.LoadSheet + "»"
                        : "в книге несколько листов «" + RestockSchema.LoadSheet + "»: " +
                          string.Join(", ", lookup.Candidates) + " - оставьте один",
                });
            }

            var sheet = lookup.Sheet;
            ExcelSheetOperations.ShowAllRows(sheet);

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Чтение листа «" + RestockSchema.LoadSheet + "»", 25);

            var layout = ReadLayout(sheet);
            var rows = ReadRows(sheet, layout);
            _logger.Information("Строк подтоварки: " + rows.Count + ".");

            if (rows.Count == 0)
            {
                throw new WarehousePlanException(
                    "На листе «" + RestockSchema.LoadSheet + "» нет ни одной строки с АЦР.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Разбор колонки «Запрет»", 30);
            var decisions = rows.Select(RestockBanRules.Decide).ToList();

            Report(progress, "Заполнение «Заметки»", 35);
            Apply(sheet, layout, rows, decisions);

            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, "Лист «" + RestockSchema.ApprovalSheet + "»", 40);
            var approvals = WriteApprovals((object)workbook, sheet, layout, rows, decisions, scope);

            // Места хранения считаются уже по «в подтоварку», урезанному разбором «Запрета».
            var storageWarnings = new List<ProcessingWarning>();
            var approvalSheet = ExcelSheetOperations.FindSheet((object)workbook, RestockSchema.ApprovalSheet, scope).Sheet;
            var storage = new ExcelRestockStorageStage(_logger, _nowProvider).Run(
                applicationObject,
                (object)workbook,
                sheet,
                approvalSheet ?? sheet,
                scope,
                (message, percent) => Report(progress, message, percent),
                cancellationToken,
                storageWarnings);

            Report(progress, "Пересчёт формул", 90);
            application.Calculation = ExcelConstants.XlCalculationAutomatic;
            application.CalculateFull();

            // Название листа читается до сохранения: после закрытия книги обращаться
            // к её листам уже нельзя.
            var result = BuildResult(
                path, ExcelSheetOperations.GetSheetName(sheet), layout, rows, decisions, approvals, storage, storageWarnings);

            Report(progress, "Сохранение файла", 95);
            workbook.Save();
            workbook.Close(true);
            closed = true;

            Report(progress, "Готово", 100);
            return result;
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

    // ================= Разметка листа =================

    private sealed record LoadLayout(HeaderMap Headers, int FirstRow, int LastRow, ColumnRange Columns)
    {
        public int Count => LastRow - FirstRow + 1;
    }

    private static LoadLayout ReadLayout(object sheet)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(sheet);
        var grid = ExcelSheetOperations.ReadBlock(
            sheet,
            bounds.FirstRow,
            Math.Min(bounds.FirstRow + HeaderResolver.DefaultScanRows - 1, bounds.LastRow),
            bounds.FirstColumn,
            bounds.LastColumn,
            withFormulas: false);

        var headers = RestockSchema.Load.ResolveHeaders(grid);
        var lastRow = ExcelSheetOperations.GetLastFilledRow(
            sheet, headers[RestockSchema.Load.Acr], bounds.LastRow);

        return new LoadLayout(
            headers,
            headers.HeaderRow + 1,
            lastRow,
            new ColumnRange(
                Math.Min(headers.Columns.Values.Min(), bounds.FirstColumn),
                Math.Max(headers.Columns.Values.Max(), bounds.LastColumn)));
    }

    /// <summary>
    /// «Прогнозный sellout» приводится к процентам: в книге колонка в процентном формате,
    /// и 80 % лежит в ячейке как 0,8. Если формат обычный, число уже в процентах.
    /// </summary>
    private static IReadOnlyList<RestockRow> ReadRows(object sheet, LoadLayout layout)
    {
        var headers = layout.Headers;
        var grid = ExcelSheetOperations.ReadBlock(
            sheet, layout.FirstRow, layout.LastRow, layout.Columns.First, layout.Columns.Last, withFormulas: false);

        var selloutColumn = headers[RestockSchema.Load.Sellout];
        var asFraction = ExcelSheetOperations.IsPercentFormat(sheet, layout.FirstRow, selloutColumn);

        var rows = new List<RestockRow>(layout.Count);
        for (var row = layout.FirstRow; row <= layout.LastRow; row++)
        {
            var sellout = grid.Number(row, selloutColumn);
            rows.Add(new RestockRow(
                row - layout.FirstRow,
                TextUtils.Normalize(grid.Text(row, headers[RestockSchema.Load.Sector])),
                TextUtils.Normalize(grid.Text(row, headers[RestockSchema.Load.Acr])),
                grid.Number(row, headers[RestockSchema.Load.Quantity]),
                grid.Value(row, headers[RestockSchema.Load.Ban]),
                grid.Value(row, headers[RestockSchema.Load.Available]),
                sellout is null ? null : asFraction ? sellout * 100d : sellout));
        }

        return rows;
    }

    // ================= Запись =================

    /// <summary>
    /// Пишет «Заметку», урезанное количество и подсветку. Пометка снимается со всей
    /// таблицы и ставится заново - иначе вчерашняя осталась бы на строке, где всё сошлось.
    ///
    /// Колонка «в подтоварку» не перекрашивается: она закрашена аналитиком (по цвету
    /// колонок потом собирают листы «из[место хранения]»), и снятие пометки стёрло бы этот цвет.
    /// </summary>
    private void Apply(
        object sheet,
        LoadLayout layout,
        IReadOnlyList<RestockRow> rows,
        IReadOnlyList<RestockDecision> decisions)
    {
        var headers = layout.Headers;
        var quantityColumn = headers[RestockSchema.Load.Quantity];

        ExcelSheetOperations.SetColumnValues(
            sheet,
            layout.FirstRow,
            headers[RestockSchema.Load.Note],
            decisions.Select(d => d.Note.Length > 0 ? (object?)d.Note : null).ToList());

        var trimmed = 0;
        foreach (var decision in decisions.Where(d => d.Quantity is not null))
        {
            ExcelSheetOperations.SetValue(
                sheet, layout.FirstRow + decision.Index, quantityColumn, decision.Quantity!.Value);
            trimmed++;
        }

        if (trimmed > 0)
        {
            _logger.Information("Урезано количество в «в подтоварку»: " + trimmed + " строк(и).");
        }

        var marked = 0;
        for (var i = 0; i < decisions.Count; i++)
        {
            var row = layout.FirstRow + i;
            var mark = decisions[i].Highlight ? RowMark.Attention : RowMark.None;

            foreach (var column in headers.Columns.Values.OrderBy(column => column))
            {
                if (column != quantityColumn)
                {
                    ExcelSheetOperations.SetRowMark(sheet, row, column, mark);
                }
            }

            if (decisions[i].Highlight)
            {
                marked++;
            }
        }

        _logger.Information("Подсвечено строк: " + marked + ".");
    }

    /// <summary>
    /// Лист согласования: те же колонки строки плюс причина. Лист пересоздаётся целиком -
    /// вчерашние строки на нём не нужны.
    /// </summary>
    private int WriteApprovals(
        object workbook,
        object loadSheet,
        LoadLayout layout,
        IReadOnlyList<RestockRow> rows,
        IReadOnlyList<RestockDecision> decisions,
        ComScope scope)
    {
        var needed = decisions.Where(d => d.NeedsApproval).ToList();

        if (ExcelSheetOperations.RemoveSheet(workbook, RestockSchema.ApprovalSheet, scope))
        {
            _logger.Information("Прошлый лист «" + RestockSchema.ApprovalSheet + "» удалён.");
        }

        if (needed.Count == 0)
        {
            return 0;
        }

        var headers = layout.Headers;
        var columns = RestockSchema.Load.ApprovalColumns
            .Where(name => headers.TryGet(name, out _))
            .ToList();

        var source = ExcelSheetOperations.ReadBlock(
            loadSheet, layout.FirstRow, layout.LastRow, layout.Columns.First, layout.Columns.Last,
            withFormulas: false);

        var block = new object?[needed.Count + 1, columns.Count + 1];
        for (var c = 0; c < columns.Count; c++)
        {
            block[0, c] = columns[c];
        }

        block[0, columns.Count] = "Почему на согласование";

        for (var i = 0; i < needed.Count; i++)
        {
            var row = layout.FirstRow + needed[i].Index;
            for (var c = 0; c < columns.Count; c++)
            {
                block[i + 1, c] = source.Value(row, headers[columns[c]]);
            }

            block[i + 1, columns.Count] = needed[i].Reason;
        }

        var sheet = ExcelSheetOperations.AddSheet(workbook, loadSheet, RestockSchema.ApprovalSheet, scope);
        ExcelSheetOperations.SetBlockValues(sheet, 1, 1, block);

        // Доля продаж на новом листе должна читаться так же, как на исходном: без формата
        // 100 % превратились бы в единицу.
        var sellout = columns.IndexOf(RestockSchema.Load.Sellout);
        if (sellout >= 0 &&
            ExcelSheetOperations.IsPercentFormat(loadSheet, layout.FirstRow, headers[RestockSchema.Load.Sellout]))
        {
            ExcelSheetOperations.SetColumnFormat(sheet, sellout + 1, "0%");
        }

        ExcelSheetOperations.AutoFitColumns(sheet, 1, columns.Count + 1);

        _logger.Information("На лист «" + RestockSchema.ApprovalSheet + "» вынесено строк: " + needed.Count + ".");
        return needed.Count;
    }

    private ProcessingResult BuildResult(
        string path,
        string sheetName,
        LoadLayout layout,
        IReadOnlyList<RestockRow> rows,
        IReadOnlyList<RestockDecision> decisions,
        int approvals,
        RestockStorageOutcome? storage,
        IReadOnlyList<ProcessingWarning> storageWarnings)
    {
        var noteColumn = ExcelColumn.ToLetters(layout.Headers[RestockSchema.Load.Note]);

        var warnings = new List<ProcessingWarning>();
        foreach (var decision in decisions.Where(d => d.Problem is not null))
        {
            var row = layout.FirstRow + decision.Index;
            warnings.Add(new ProcessingWarning(
                "АЦР " + rows[decision.Index].Acr + ": " + decision.Problem,
                RestockSchema.LoadSheet + ", строка " + row,
                sheetName,
                noteColumn + row));
        }

        var ok = decisions.Count(d => d.Note == RestockSchema.NoteOk);
        var approve = decisions.Count(d => d.Note == RestockSchema.NoteApprove);
        var undecided = decisions.Count(d => d.Note.Length == 0);
        var highlighted = decisions.Count(d => d.Highlight);

        var counters = new List<ProcessingCounter>
        {
            new("строк подтоварки", rows.Count),
            new("«Ок»", ok),
            new("«Согласовать»", approve, approve > 0),
            new("на согласование", approvals, approvals > 0),
            new("подсвечено строк", highlighted, highlighted > 0),
            new("без решения", undecided, undecided > 0),
        };

        if (storage is not null)
        {
            counters.Add(new ProcessingCounter("строк с местом хранения", storage.Placed));
            counters.Add(new ProcessingCounter("не собрано", storage.NotCollected, storage.NotCollected > 0));
            counters.Add(new ProcessingCounter("загрузочников", storage.Loaders, storage.Loaders > 0));
            counters.Add(new ProcessingCounter("шт. в загрузочниках", storage.LoaderQuantity));
        }

        warnings.AddRange(storageWarnings);
        return new ProcessingResult(path, counters, warnings);
    }

    private static void Report(IProgress<ProcessingStage>? progress, string message, int percent) =>
        progress?.Report(new ProcessingStage(message, percent));
}
