using System.Globalization;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>Что получилось на втором этапе: счётчики для карточки результата и замечания.</summary>
internal sealed record DistributionOutcome(
    int StockColumns,
    int Blocks,
    int ZeroedHubMinimums,
    IReadOnlyList<ProcessingWarning> Warnings);

/// <summary>
/// Второй этап распреда: лист «остатки» и лист «Распред».
///
/// «остатки» - это транспонированная выгрузка «Остатки Н», отобранная по сектору поставки.
/// «Распред» почти целиком состоит из формул, которые сами подстраиваются под номер своей
/// колонки, поэтому новый блок АЦР - это копия первого блока, а не набор новых формул.
/// </summary>
internal sealed class ExcelDistributionStage
{
    private const int HeaderScanRows = 15;
    private const int ChunkRows = 20000;

    /// <summary>
    /// Насколько далеко друг от друга могут стоять строки сектора, чтобы читать их
    /// одним куском. Выгрузка обычно отсортирована по сектору, и тогда чтение одно.
    /// </summary>
    private const int RunGap = 50;

    /// <summary>Рамка, в которую вписывается фотография в заголовке блока АЦР.</summary>
    private const double PhotoWidth = 150d;

    private const double PhotoHeight = 110d;

    private readonly IAppLogger _logger;
    private readonly Func<DateTime> _nowProvider;

    public ExcelDistributionStage(IAppLogger logger, Func<DateTime> nowProvider)
    {
        _logger = logger;
        _nowProvider = nowProvider;
    }

    public DistributionOutcome Run(
        object application,
        object stockSourceSheet,
        object stockSheet,
        object distributionSheet,
        object seasonalitySheet,
        object pricesSheet,
        int pricesFirstRow,
        ColumnRange pricesColumns,
        IReadOnlyList<PriceRowValues> priceRows,
        CancellationToken cancellationToken)
    {
        var warnings = new List<ProcessingWarning>();

        var sector = ResolveSector(priceRows, warnings);
        var stock = BuildStock(stockSourceSheet, sector, warnings);
        WriteStockSheet(stockSheet, stock);

        cancellationToken.ThrowIfCancellationRequested();

        var layout = ReadLayout(distributionSheet);
        var blocks = Resize(application, distributionSheet, layout, priceRows.Count, warnings);

        // Разметка перечитывается ради сдвинувшихся адресов, но число блоков берётся
        // задуманное: вставленные колонки пока пустые, и по листу их не сосчитать.
        layout = ReadLayout(distributionSheet).WithBlockCount(blocks);

        RetargetStockReferences(distributionSheet, layout, stock);

        // Первый блок - образец для остальных. Если в прошлый раз у него обнулили
        // «мин запас на Хаб», ноль разошёлся бы по всем блокам; формула возвращается до копирования.
        RestoreHubMinimums(distributionSheet, layout);
        CopyFirstBlock(application, distributionSheet, layout);
        CopyPhotos(
            application, distributionSheet, layout, pricesSheet, pricesFirstRow, pricesColumns,
            priceRows.Count, warnings);

        if (sector.Length > 0)
        {
            ExcelSheetOperations.SetValue(
                distributionSheet, layout.SectorCell.Row, layout.SectorCell.Column, sector);
            _logger.Information("Сектор «" + sector + "» записан в " + layout.SectorCell + ".");
        }

        WriteSeasonality(distributionSheet, layout, seasonalitySheet, priceRows, sector, warnings);

        return new DistributionOutcome(stock.ColumnCount, blocks, 0, warnings);
    }

    /// <summary>
    /// Пересчёт только остатков: лист «остатки» собирается заново из «Остатков Н»,
    /// ссылки «Распреда» на его последнюю колонку переставляются. Блоки, фотографии,
    /// сезонность и «Загрузочник» не трогаются. «Мин запас на Хаб» возвращается к общей
    /// настройке, чтобы правило 30 % проверилось заново по свежим остаткам: это делает
    /// <see cref="Finish"/> после пересчёта формул.
    /// </summary>
    public DistributionOutcome RunStockOnly(
        object stockSourceSheet,
        object stockSheet,
        object distributionSheet,
        IReadOnlyList<PriceRowValues> priceRows)
    {
        var warnings = new List<ProcessingWarning>();

        var sector = ResolveSector(priceRows, warnings);
        var stock = BuildStock(stockSourceSheet, sector, warnings);
        WriteStockSheet(stockSheet, stock);

        var layout = ReadLayout(distributionSheet);
        RetargetStockReferences(distributionSheet, layout, stock);
        RestoreHubMinimums(distributionSheet, layout);

        return new DistributionOutcome(stock.ColumnCount, layout.BlockCount, 0, warnings);
    }

    /// <summary>
    /// Возвращает «мин запас на Хаб» к общей настройке там, где его раньше обнулила программа.
    /// Формула берётся у блока, где она сохранилась (=$L$6 - ссылка абсолютная, годится
    /// для любого блока). Меняются только нули: другое число вписал человек, его не трогаем.
    /// </summary>
    private void RestoreHubMinimums(object sheet, DistributionLayout layout)
    {
        if (layout.BlockCount == 0)
        {
            return;
        }

        var columns = Enumerable.Range(0, layout.BlockCount).Select(layout.BlockColumn).ToList();
        var formulas = columns
            .Select(column => ExcelSheetOperations.GetFormula(sheet, layout.HubMinimumRow, column))
            .ToList();

        var template = formulas.FirstOrDefault(formula => formula is not null && formula.StartsWith('='));
        if (template is null)
        {
            return;
        }

        var restored = 0;
        for (var i = 0; i < columns.Count; i++)
        {
            if (formulas[i] is { } formula && formula.StartsWith('='))
            {
                continue;
            }

            var value = TextUtils.CellToDouble(ExcelSheetOperations.GetValue(sheet, layout.HubMinimumRow, columns[i]));
            if (value is not 0d)
            {
                continue;
            }

            ExcelSheetOperations.SetFormula(sheet, layout.HubMinimumRow, columns[i], template);
            restored++;
        }

        if (restored > 0)
        {
            _logger.Information("«Мин запас на Хаб» возвращён к общей настройке в " + restored + " блок(ах): " + template + ".");
        }
    }

    /// <summary>
    /// Шаги, которым нужны посчитанные значения: обнуление «мин запаса на Хаб» там, где
    /// остаток вышел меньше 30 %, и сортировка таблицы РТТ по рейтингу.
    /// Вызывается после пересчёта формул; после обнуления нужен ещё один пересчёт.
    /// </summary>
    public DistributionOutcome Finish(
        object sheet,
        DistributionOutcome outcome,
        Action recalculate)
    {
        var layout = ReadLayout(sheet);
        var warnings = outcome.Warnings.ToList();

        var zeroed = ApplyHubMinimums(sheet, layout, warnings);
        if (zeroed > 0)
        {
            recalculate();
        }

        SortByRating(sheet, layout, warnings);

        return outcome with { ZeroedHubMinimums = zeroed, Warnings = warnings };
    }

    /// <summary>
    /// Где доля остатка меньше 30 %, «мин запас на Хаб» обнуляется: хаб придерживает
    /// меньше, в загрузку уходит меньше, и остаток подрастает. Формула =$L$6 при этом
    /// заменяется числом - для этого АЦР общая настройка больше не действует.
    /// </summary>
    private int ApplyHubMinimums(object sheet, DistributionLayout layout, List<ProcessingWarning> warnings)
    {
        var zeroed = 0;

        // Название листа берётся из книги, а не из схемы: в схеме записано начало
        // названия, а перейти нужно на тот лист, который в книге на самом деле.
        var sheetName = ExcelSheetOperations.GetSheetName(sheet);

        for (var index = 0; index < layout.BlockCount; index++)
        {
            var column = layout.BlockColumn(index);
            var raw = ExcelSheetOperations.GetValue(sheet, layout.RemainderRow, column + 1);
            var codeValue = ExcelSheetOperations.GetValue(sheet, layout.CodeRow, column);

            // Код ошибки Excel отдаёт большим отрицательным числом. Без этой проверки
            // «#Н/Д» прошло бы как «остаток меньше 30 %» и обнулило запас ни за что.
            if (CellError.IsError(raw))
            {
                warnings.Add(new ProcessingWarning(
                    "Доля остатка не посчиталась (в ячейке ошибка формулы) - «мин запас на Хаб» " +
                    "оставлен как был. Проверьте этот АЦР вручную.",
                    PriceSchema.DistributionSheet + ", " +
                    new CellRef(layout.RemainderRow, column + 1),
                    sheetName,
                    new CellRef(layout.RemainderRow, column + 1).ToString()));
                continue;
            }

            var share = TextUtils.CellToDouble(raw);
            if (share is null || share.Value >= PriceSchema.Distribution.MinimumRemainderShare)
            {
                continue;
            }

            var code = CellError.IsError(codeValue)
                ? "без кода"
                : TextUtils.CellToString(codeValue);

            ExcelSheetOperations.SetValue(sheet, layout.HubMinimumRow, column, 0d);
            zeroed++;

            warnings.Add(new ProcessingWarning(
                "У кода " + code + " остаток вышел " + (share.Value * 100d).ToString("0.#") +
                " %, меньше 30 %: «мин запас на Хаб» обнулён.",
                PriceSchema.DistributionSheet + ", " + new CellRef(layout.HubMinimumRow, column),
                sheetName,
                new CellRef(layout.HubMinimumRow, column).ToString()));
        }

        if (zeroed > 0)
        {
            _logger.Information("Обнулено «мин запас на Хаб»: " + zeroed + ".");
        }

        return zeroed;
    }

    /// <summary>
    /// Сортировка таблицы РТТ по колонке «рейтинг». Формулы строк ссылаются только на свою
    /// же строку и на целые колонки, поэтому перестановка их не ломает.
    /// </summary>
    private void SortByRating(object sheet, DistributionLayout layout, List<ProcessingWarning> warnings)
    {
        var grid = ExcelSheetOperations.ReadBlock(
            sheet, layout.HeaderRow, layout.HeaderRow, 1, layout.LastBlockColumn, withFormulas: false);

        int? ratingColumn = null;
        for (var column = 1; column <= layout.LastBlockColumn; column++)
        {
            if (TextUtils.EqualsKey(grid.Text(layout.HeaderRow, column), "рейтинг"))
            {
                ratingColumn = column;
                break;
            }
        }

        if (ratingColumn is null)
        {
            warnings.Add(new ProcessingWarning(
                "В строке заголовков листа «" + PriceSchema.DistributionSheet +
                "» нет колонки «рейтинг»: таблица РТТ не отсортирована.",
                PriceSchema.DistributionSheet));
            return;
        }

        ExcelSheetOperations.SortRange(
            sheet,
            layout.StoreFirstRow,
            layout.StoreLastRow,
            1,
            layout.LastBlockColumn,
            ratingColumn.Value);

        _logger.Information(
            "Таблица РТТ отсортирована по колонке " + ExcelColumn.ToLetters(ratingColumn.Value) + ".");
    }

    // ================= Сектор =================

    /// <summary>
    /// Сектор поставки. Обычно он один на всю поставку; если в «Цены» их несколько,
    /// программа берёт самый частый и пишет замечание - выбор сектора меняет всю таблицу РТТ.
    /// </summary>
    private string ResolveSector(IReadOnlyList<PriceRowValues> rows, List<ProcessingWarning> warnings)
    {
        var sectors = rows
            .Select(row => TextUtils.Normalize(row.Reference?.Sector))
            .Where(sector => sector.Length > 0)
            .GroupBy(sector => sector, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Count())
            .ToList();

        if (sectors.Count == 0)
        {
            warnings.Add(new ProcessingWarning(
                "Ни у одной строки «Цены» не заполнен сектор: таблица РТТ на листе «" +
                PriceSchema.DistributionSheet + "» не перенастроена.",
                PriceSchema.PricesSheet));
            return string.Empty;
        }

        if (sectors.Count > 1)
        {
            warnings.Add(new ProcessingWarning(
                "В поставке несколько секторов (" +
                string.Join(", ", sectors.Select(g => g.Key + " - " + g.Count() + " шт.")) +
                "). Взят самый частый: «" + sectors[0].Key + "».",
                PriceSchema.PricesSheet));
        }

        return sectors[0].Key;
    }

    // ================= Лист «остатки» =================

    private StockSheetContent BuildStock(object sheet, string sector, List<ProcessingWarning> warnings)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(sheet);
        var headerGrid = ExcelSheetOperations.ReadBlock(
            sheet,
            bounds.FirstRow,
            Math.Min(bounds.FirstRow + HeaderScanRows - 1, bounds.LastRow),
            bounds.FirstColumn,
            bounds.LastColumn,
            withFormulas: false);

        var headers = HeaderResolver.Resolve(
            headerGrid, PriceSchema.StockSourceSheet, PriceSchema.StockSource.Specs);

        var columns = StockSheetBuilder.ChooseColumns(headerGrid, headers.HeaderRow, out var labels);

        var sectorColumn = headers[PriceSchema.StockSource.Sector];
        var lastRow = Math.Min(
            bounds.LastRow,
            ExcelSheetOperations.GetLastFilledRow(sheet, sectorColumn, bounds.LastRow));

        var matching = FindSectorRows(sheet, sectorColumn, headers.HeaderRow + 1, lastRow, sector);
        _logger.Information(
            "«" + PriceSchema.StockSourceSheet + "»: строк сектора «" + sector + "» - " + matching.Count +
            " из " + (lastRow - headers.HeaderRow) + ".");

        if (matching.Count == 0)
        {
            warnings.Add(new ProcessingWarning(
                "На листе «" + PriceSchema.StockSourceSheet + "» нет ни одной строки сектора «" +
                sector + "»: лист «" + PriceSchema.StockSheet + "» очищен.",
                PriceSchema.StockSourceSheet));
        }

        var articleColumn = headers[PriceSchema.StockSource.Article];
        var colorColumn = headers[PriceSchema.StockSource.Color];
        var sizeColumn = headers[PriceSchema.StockSource.Size];

        var rows = new List<StockSourceRow>(matching.Count);
        foreach (var (first, last) in BuildRuns(matching))
        {
            var chunk = ExcelSheetOperations.ReadBlock(
                sheet, first, last, bounds.FirstColumn, bounds.LastColumn, withFormulas: false);

            foreach (var row in matching.Where(r => r >= first && r <= last))
            {
                var acr = StockSheetBuilder.BuildAcr(
                    chunk.Text(row, articleColumn),
                    chunk.Text(row, colorColumn),
                    chunk.Text(row, sizeColumn));

                var values = new List<object?>(columns.Count);
                foreach (var column in columns)
                {
                    values.Add(column == 0 ? acr : chunk.Value(row, column));
                }

                rows.Add(new StockSourceRow(acr, values));
            }
        }

        return new StockSheetContent(labels, rows);
    }

    /// <summary>Номера строк нужного сектора. Читается только одна колонка.</summary>
    private static List<int> FindSectorRows(
        object sheet, int sectorColumn, int firstRow, int lastRow, string sector)
    {
        var wanted = TextUtils.NormalizeKey(sector);
        var result = new List<int>();

        for (var start = firstRow; start <= lastRow; start += ChunkRows)
        {
            var end = Math.Min(start + ChunkRows - 1, lastRow);
            var chunk = ExcelSheetOperations.ReadBlock(
                sheet, start, end, sectorColumn, sectorColumn, withFormulas: false);

            for (var row = start; row <= end; row++)
            {
                if (string.Equals(
                        TextUtils.NormalizeKey(chunk.Text(row, sectorColumn)),
                        wanted,
                        StringComparison.Ordinal))
                {
                    result.Add(row);
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Склеивает номера строк в непрерывные куски: выгрузка отсортирована по сектору,
    /// и тогда весь сектор читается одним обращением вместо сотен построчных.
    /// </summary>
    private static List<(int First, int Last)> BuildRuns(IReadOnlyList<int> rows)
    {
        var runs = new List<(int First, int Last)>();
        if (rows.Count == 0)
        {
            return runs;
        }

        var first = rows[0];
        var previous = rows[0];

        foreach (var row in rows.Skip(1))
        {
            if (row - previous <= RunGap)
            {
                previous = row;
                continue;
            }

            runs.Add((first, previous));
            first = row;
            previous = row;
        }

        runs.Add((first, previous));
        return runs;
    }

    /// <summary>
    /// Пишет лист «остатки»: в первой колонке дата и коды магазинов, во второй подписи
    /// строк, дальше по колонке на каждый АЦР.
    /// </summary>
    private void WriteStockSheet(object sheet, StockSheetContent stock)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(sheet);
        ExcelSheetOperations.ClearRange(
            sheet, bounds.FirstRow, bounds.LastRow, bounds.FirstColumn, bounds.LastColumn);

        if (stock.RowCount == 0)
        {
            return;
        }

        var block = new object?[stock.RowCount, stock.ColumnCount + 1];
        for (var row = 0; row < stock.RowCount; row++)
        {
            block[row, 0] = stock.Labels[row];
            for (var column = 0; column < stock.ColumnCount; column++)
            {
                block[row, column + 1] = stock.Rows[column].Values[row];
            }
        }

        // Строки, которые Excel принял бы за числа, переводятся в текстовый формат до записи:
        // иначе перечень магазинов «150,155,156,…» станет числом 1,5E+113. Формат ставится
        // построчно - строка листа это одна колонка выгрузки, и тип значений в ней один.
        var textRows = new List<int>();
        for (var row = 0; row < stock.RowCount; row++)
        {
            var risky = false;
            for (var column = 0; column < stock.ColumnCount && !risky; column++)
            {
                risky = block[row, column + 1] is string text && CellError.LooksNumericToExcel(text);
            }

            if (risky)
            {
                textRows.Add(row + 1);
                ExcelSheetOperations.SetRangeAsText(sheet, row + 1, row + 1, 2, stock.ColumnCount + 2);
            }
        }

        ExcelSheetOperations.SetBlockValues(sheet, 1, 2, block);

        if (textRows.Count > 0)
        {
            _logger.Information(
                "Строк «" + PriceSchema.StockSheet + "», записанных как текст: " +
                string.Join(", ", textRows.Select(row => stock.Labels[row - 1])) + ".");
        }

        // Дата поставки - в первой ячейке листа: по ней видно, за какой день остатки.
        ExcelSheetOperations.SetValue(sheet, 1, 1, _nowProvider().Date);

        // Код магазина рядом со строками «тз» и «в_пути»: по нему «Распред» и собирает
        // остаток магазина. У строк «прод» кода нет - проданное остатком не является.
        // Первая строка пропускается: там дата.
        if (stock.RowCount > 1)
        {
            var codes = new object?[stock.RowCount - 1, 1];
            for (var row = 1; row < stock.RowCount; row++)
            {
                codes[row - 1, 0] = stock.IsStockRow(row)
                    ? "=LEFT(B" + (row + 1).ToString(CultureInfo.InvariantCulture) + ",3)"
                    : null;
            }

            ExcelSheetOperations.SetBlockFormulas(sheet, 2, 1, codes);
        }

        _logger.Information(
            "Лист «" + PriceSchema.StockSheet + "» перезаполнен: " + stock.RowCount + " строк, " +
            stock.ColumnCount + " АЦР.");
    }

    // ================= Лист «Распред» =================

    private static DistributionLayout ReadLayout(object sheet) =>
        DistributionLayout.Read(ExcelSheetOperations.ReadGrid(sheet, withFormulas: true));

    /// <summary>
    /// Доводит число блоков до числа кодов. Колонки вставляются и удаляются внутри
    /// занятого блоками диапазона: тогда Excel сам растягивает формулы, которые охватывают
    /// все блоки сразу - «Итого» строки РТТ и суммы по строкам «Заказ», «Ед», «МП».
    /// </summary>
    private int Resize(
        object application,
        object sheet,
        DistributionLayout layout,
        int wanted,
        List<ProcessingWarning> warnings)
    {
        if (wanted <= 0)
        {
            warnings.Add(new ProcessingWarning(
                "В поставке нет ни одного кода: блоки на листе «" + PriceSchema.DistributionSheet +
                "» не перестроены.",
                PriceSchema.DistributionSheet));
            return layout.BlockCount;
        }

        var difference = wanted - layout.BlockCount;
        if (difference == 0)
        {
            return wanted;
        }

        var width = PriceSchema.Distribution.BlockWidth;
        if (difference > 0)
        {
            // Когда блок на листе один, вставка приходится внутрь него самого,
            // поэтому образец на время убирается в сторону.
            var aside = layout.BlockCount == 1 ? MoveTemplateAside(application, sheet, layout) : 0;

            ExcelSheetOperations.InsertColumns(sheet, layout.LastBlockColumn, difference * width);

            if (aside > 0)
            {
                RestoreTemplate(application, sheet, layout, aside + (difference * width));
            }

            _logger.Information("На листе «Распред» добавлено блоков: " + difference + ".");
        }
        else
        {
            ExcelSheetOperations.DeleteColumns(
                sheet, layout.BlockColumn(wanted), -difference * width);
            _logger.Information("На листе «Распред» удалено блоков: " + -difference + ".");
        }

        return wanted;
    }

    /// <summary>
    /// Откладывает блок-образец правее всего занятого и возвращает его колонку.
    ///
    /// Новые колонки вставляются перед последней занятой - только так Excel сам растягивает
    /// формулы, которые охватывают все блоки сразу: «Итого» в строках РТТ и суммы «Заказ»,
    /// «Ед», «МП». Пока блоков несколько, под вставку попадает последний блок, а образцом
    /// служит первый, и размножение всё восстанавливает. Но когда блок на листе один, он же
    /// и образец: его «в загрузку» уехала бы вправо, на её месте осталась бы пустая колонка,
    /// и размножение раздало бы эту пустоту всем блокам.
    ///
    /// Копия и копия обратно дают ровно исходные формулы: ссылки внутри блока смещаются
    /// туда и обратно на одно и то же число колонок, а всё, на что блок смотрит снаружи,
    /// записано в нём абсолютными адресами.
    /// </summary>
    private int MoveTemplateAside(object application, object sheet, DistributionLayout layout)
    {
        var width = PriceSchema.Distribution.BlockWidth;
        var aside = Math.Max(
            ExcelSheetOperations.GetUsedBounds(sheet).LastColumn, layout.LastBlockColumn) + 2;

        ExcelSheetOperations.CopyRange(
            application,
            sheet,
            1,
            layout.FirstBlockColumn,
            layout.StoreLastRow,
            layout.FirstBlockColumn + width - 1,
            1,
            aside);

        return aside;
    }

    /// <summary>Возвращает отложенный блок-образец на место и убирает копию.</summary>
    private void RestoreTemplate(object application, object sheet, DistributionLayout layout, int aside)
    {
        var width = PriceSchema.Distribution.BlockWidth;

        ExcelSheetOperations.CopyRange(
            application,
            sheet,
            1,
            aside,
            layout.StoreLastRow,
            aside + width - 1,
            1,
            layout.FirstBlockColumn);

        ExcelSheetOperations.DeleteColumns(sheet, aside, width);

        _logger.Information(
            "На листе «Распред» блок-образец был единственным: вставка пришлась внутрь него, " +
            "образец восстановлен из копии в " + ExcelColumn.ToLetters(aside) + ".");
    }

    /// <summary>
    /// Переводит ссылки на лист «остатки» на его фактические размеры. Формулы написаны
    /// под прошлую выгрузку, а число АЦР и число строк меняются с каждой поставкой.
    /// </summary>
    private void RetargetStockReferences(object sheet, DistributionLayout layout, StockSheetContent stock)
    {
        if (stock.RowCount == 0 || stock.ColumnCount == 0)
        {
            return;
        }

        var lastColumn = 2 + stock.ColumnCount;
        var lastRow = stock.RowCount;

        for (var row = layout.HeaderRow - 1; row <= layout.StoreLastRow; row++)
        {
            for (var column = layout.FirstBlockColumn;
                 column < layout.FirstBlockColumn + PriceSchema.Distribution.BlockWidth;
                 column++)
            {
                var formula = ExcelSheetOperations.GetFormula(sheet, row, column);
                var repaired = StockReferenceRepair.Retarget(
                    formula, PriceSchema.StockSheet, lastRow, lastColumn);

                if (repaired is not null)
                {
                    ExcelSheetOperations.SetFormula(sheet, row, column, repaired);
                    _logger.Debug(
                        "Ссылка на «остатки» в " + new CellRef(row, column) + " приведена к виду " +
                        repaired + ".");
                }
            }
        }
    }

    /// <summary>
    /// Размножает первый блок по остальным. Формулы блока опираются на номер своей колонки
    /// (ИНДЕКС(Цены!$D:$D; 2+(СТОЛБЕЦ()-СТОЛБЕЦ($N$18))/4)), поэтому копия сама берёт
    /// нужный код - переписывать формулы не нужно.
    /// </summary>
    private void CopyFirstBlock(object application, object sheet, DistributionLayout layout)
    {
        if (layout.BlockCount <= 1)
        {
            return;
        }

        var width = PriceSchema.Distribution.BlockWidth;
        var firstRow = 1;
        var lastRow = layout.StoreLastRow;

        for (var index = 1; index < layout.BlockCount; index++)
        {
            ExcelSheetOperations.CopyRange(
                application,
                sheet,
                firstRow,
                layout.FirstBlockColumn,
                lastRow,
                layout.FirstBlockColumn + width - 1,
                firstRow,
                layout.BlockColumn(index));
        }

        _logger.Information("Первый блок размножен на " + (layout.BlockCount - 1) + " повтор(ов).");
    }

    /// <summary>
    /// Ставит в заголовок каждого блока фотографию его АЦР - ту, что лежит на листе «Цены»
    /// в строке этого товара.
    ///
    /// Своей картинки у блока не появляется само собой: размножение первого блока копирует
    /// картинку образца, и во всех блоках оказывается одна и та же. Поэтому картинки
    /// из области блоков сначала удаляются, а потом ставятся заново - по строке на блок.
    /// </summary>
    private void CopyPhotos(
        object application,
        object sheet,
        DistributionLayout layout,
        object pricesSheet,
        int pricesFirstRow,
        ColumnRange pricesColumns,
        int rowCount,
        List<ProcessingWarning> warnings)
    {
        using var scope = new ComScope();

        var photos = ExcelPictures.ByRow(
            pricesSheet,
            new SheetArea(
                pricesFirstRow, pricesFirstRow + rowCount - 1, pricesColumns.First, pricesColumns.Last),
            scope);

        // Убирать старые картинки нужно в любом случае, даже когда ставить нечего:
        // размножение первого блока только что раздало картинку образца по всем блокам,
        // и без этого шага во всех блоках осталась бы одна и та же фотография.
        var removed = ExcelPictures.DeleteIn(
            sheet,
            new SheetArea(
                1,
                layout.HeaderRow - 1,
                layout.FirstBlockColumn,
                layout.BlockColumn(layout.BlockCount - 1) + PriceSchema.Distribution.BlockWidth - 1),
            overlapping: true);

        if (photos.Count == 0)
        {
            _logger.Information(
                "Фотографии в заголовках блоков: убрано " + removed + ", ставить нечего.");
            warnings.Add(new ProcessingWarning(
                "На листе «" + PriceSchema.PricesSheet + "» нет фотографий товара: в заголовках блоков " +
                "листа «" + PriceSchema.DistributionSheet + "» их тоже не будет. Фотографии переносит " +
                "подготовка распреда - если их нет, значит их не было и в инвойсе.",
                PriceSchema.PricesSheet));
            return;
        }

        var copied = 0;
        for (var index = 0; index < layout.BlockCount; index++)
        {
            if (!photos.TryGetValue(pricesFirstRow + index, out var shape))
            {
                continue;
            }

            // Картинка стоит во второй колонке блока - так она не закрывает код АЦР.
            if (ExcelPictures.CopyTo(
                    application, shape, sheet, 1, layout.BlockColumn(index) + 1,
                    PhotoWidth, PhotoHeight, fillFrame: true, _logger))
            {
                copied++;
            }
        }

        _logger.Information(
            "Фотографии в заголовках блоков: убрано " + removed + ", поставлено " + copied +
            " на " + layout.BlockCount + " блок(ов).");

        if (copied < layout.BlockCount)
        {
            warnings.Add(new ProcessingWarning(
                "Фотографии проставлены не во всех блоках листа «" + PriceSchema.DistributionSheet +
                "»: " + copied + " из " + layout.BlockCount +
                ". У остальных АЦР нет фотографии на листе «" + PriceSchema.PricesSheet + "».",
                PriceSchema.DistributionSheet));
        }
    }

    /// <summary>Подгруппы, месяцы и доли сезонности.</summary>
    private void WriteSeasonality(
        object sheet,
        DistributionLayout layout,
        object seasonalitySheet,
        IReadOnlyList<PriceRowValues> priceRows,
        string sector,
        List<ProcessingWarning> warnings)
    {
        var subgroups = priceRows
            .Where(row => row.Subgroup.Length > 0)
            .Select(row => new SubgroupKey(row.Reference?.Group ?? string.Empty, row.Subgroup))
            .DistinctBy(key => TextUtils.NormalizeKey(key.Group) + "|" + TextUtils.NormalizeKey(key.Subgroup))
            .OrderBy(key => key.Subgroup, StringComparer.CurrentCulture)
            .ToList();

        if (subgroups.Count == 0)
        {
            warnings.Add(new ProcessingWarning(
                "Ни у одной строки «Цены» не заполнена подгруппа: блок сезонности не заполнен.",
                PriceSchema.PricesSheet));
            return;
        }

        var months = MonthCalendar.SeasonFrom(_nowProvider());
        warnings.Add(new ProcessingWarning(
            "Сезонность посчитана за " + string.Join(", ", months.Select(month => month.ToString())) +
            ": три месяца, начиная со следующего за текущим. В расчёт идут недели, у которых " +
            "в месяце не меньше " + MonthCalendar.MinimumDaysInMonth + " дней.",
            PriceSchema.DistributionSheet));

        var seasonality = SeasonalitySheetReader.Read(
            ExcelSheetOperations.ReadGrid(seasonalitySheet, withFormulas: false));

        var plan = SeasonalityCalculator.Build(subgroups, months, seasonality);
        warnings.AddRange(plan.Warnings);

        // Для трёх секторов доли пересчитываются по шаблону: доля месяца делится на сумму
        // долей за месяцы, которые идут в расчёт. Остальным секторам пересчёт не делают.
        var recalcMonths = SeasonalityRecalc.MonthsFor(sector);
        var shares = plan.Rows
            .Select(row => SeasonalityRecalc.Apply(row.Shares, recalcMonths))
            .ToList();

        if (recalcMonths > 0)
        {
            _logger.Information(
                "Сектор «" + sector + "»: доли сезонности пересчитаны по " + recalcMonths + " месяцам.");
            warnings.Add(new ProcessingWarning(
                "Сектор «" + sector + "»: доли сезонности пересчитаны по " + recalcMonths +
                " месяцам - доля месяца делится на их сумму. " +
                (recalcMonths < layout.MonthCount
                    ? "Доли остальных месяцев оставлены такими, как посчитались с «" +
                      PriceSchema.SeasonalitySheet + "»."
                    : "Так задано шаблоном пересчёта."),
                PriceSchema.DistributionSheet));
        }

        var capacity = layout.SubgroupCapacity;
        var written = Math.Min(plan.Rows.Count, capacity);
        if (plan.Rows.Count > capacity)
        {
            warnings.Add(new ProcessingWarning(
                "Подгрупп в поставке " + plan.Rows.Count + ", а в блоке сезонности помещается " +
                capacity + ": лишние не записаны. Формулы распределения смотрят только в этот блок.",
                PriceSchema.DistributionSheet));
        }

        // «Мин раздача» вносит человек по подгруппе. Прошлые значения стояли напротив прошлого
        // списка подгрупп: остаются только у совпавших подгрупп, остальное убирается.
        var minimumColumn = layout.Seasonality.LastColumn;
        var old = ExcelSheetOperations.ReadBlock(
            sheet,
            layout.Seasonality.FirstRow,
            layout.Seasonality.LastRow,
            layout.SubgroupColumn,
            minimumColumn,
            withFormulas: false);
        var previous = Enumerable.Range(layout.Seasonality.FirstRow, layout.Seasonality.RowCount)
            .Select(row => ((string?)old.Text(row, layout.SubgroupColumn), old.Value(row, minimumColumn)))
            .ToList();
        var minimums = SeasonalityMinimums.Carry(
            previous, plan.Rows.Take(written).Select(row => row.Key.Subgroup).ToList());

        // Старый список очищается целиком, вместе с «Мин раздачей»: подгрупп могло быть больше, чем сейчас.
        ExcelSheetOperations.ClearRange(
            sheet,
            layout.Seasonality.FirstRow,
            layout.Seasonality.LastRow,
            layout.SubgroupColumn,
            minimumColumn);

        var monthNames = new object?[1, layout.MonthCount];
        for (var i = 0; i < layout.MonthCount; i++)
        {
            monthNames[0, i] = i < months.Count ? months[i].Name : null;
        }

        ExcelSheetOperations.SetBlockValues(sheet, layout.MonthRow, layout.MonthColumn(0), monthNames);

        if (written > 0)
        {
            var block = new object?[written, 2 + layout.MonthCount];
            for (var row = 0; row < written; row++)
            {
                block[row, 0] = plan.Rows[row].Key.Subgroup;
                for (var month = 0; month < layout.MonthCount; month++)
                {
                    block[row, month + 1] = month < shares[row].Count ? shares[row][month] : null;
                }

                block[row, 1 + layout.MonthCount] = minimums[row];
            }

            ExcelSheetOperations.SetBlockValues(
                sheet, layout.Seasonality.FirstRow, layout.SubgroupColumn, block);
        }

        _logger.Information(
            "Блок сезонности: месяцы " + string.Join(", ", months.Select(m => m.Name)) +
            ", подгрупп " + written + ", «Мин раздача» сохранена у " + minimums.Count(value => value is not null) + ".");
    }

}
