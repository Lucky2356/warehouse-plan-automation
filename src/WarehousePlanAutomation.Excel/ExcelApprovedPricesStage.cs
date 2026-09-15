using System.Runtime.InteropServices;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>Сколько согласованных цен проставлено и что не нашлось.</summary>
internal sealed record ApprovedPricesOutcome(int Filled, int Missing, IReadOnlyList<ProcessingWarning> Warnings);

/// <summary>
/// «Согласованные цены» из файлов согласования.
///
/// Вручную в колонку протягивают ВПР на файл из папки КМ «Согласование цен» → сезон,
/// найденный по номеру поставки. Программа делает то же, но пишет значения: формула
/// со ссылкой на закрытый файл на сетевом диске теряет цифры при первом пересчёте без
/// доступа к папке, а значения остаются.
///
/// Папку указывает человек, и только в ней программа ищет. Файлы открываются только
/// для чтения и закрываются без сохранения - чужой файл остаётся нетронутым.
/// </summary>
internal sealed class ExcelApprovedPricesStage
{
    /// <summary>Вложенность папок под выбранной: папка КМ → «Согласование цен» → сезон.</summary>
    private const int FolderDepth = 3;

    /// <summary>msoAutomationSecurityForceDisable: макросы чужого файла не запускаются.</summary>
    private const int AutomationSecurityForceDisable = 3;

    private readonly IAppLogger _logger;

    /// <summary>Прочитанные файлы: один и тот же файл бывает ближайшим сразу для нескольких поставок.</summary>
    private readonly Dictionary<string, IReadOnlyDictionary<string, double>?> _tables =
        new(StringComparer.OrdinalIgnoreCase);

    public ExcelApprovedPricesStage(IAppLogger logger)
    {
        _logger = logger;
    }

    public ApprovedPricesOutcome Run(
        object application,
        string source,
        object pricesSheet,
        int firstDataRow,
        int priceColumn,
        IReadOnlyList<PriceRowValues> rows,
        CancellationToken cancellationToken)
    {
        var warnings = new List<ProcessingWarning>();
        var sheetName = ExcelSheetOperations.GetSheetName(pricesSheet);

        // Выбран файл - цены сначала берутся из него, а недостающие ищутся в его папке.
        var chosenFile = File.Exists(source) ? Path.GetFullPath(source) : null;
        var folder = chosenFile is null ? source : Path.GetDirectoryName(chosenFile) ?? source;

        if (!Directory.Exists(folder))
        {
            warnings.Add(new ProcessingWarning(
                "Файл или папка согласования цен недоступны: " + source + ". «" + PriceSchema.Prices.ApprovedPrice +
                "» не заполнены программой и остались такими, как были в книге.",
                PriceSchema.PricesSheet));
            return new ApprovedPricesOutcome(0, rows.Count, warnings);
        }

        var files = ListFiles(folder);
        _logger.Information("Папка согласования цен " + folder + ": файлов Excel " + files.Count + ".");

        var found = new Dictionary<int, double>();
        var missing = 0;

        foreach (var supply in rows
                     .GroupBy(row => TextUtils.Normalize(row.Supply), StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var supplyRows = supply.Where(row => row.Barcode.Length > 0).ToList();
            if (supplyRows.Count == 0)
            {
                continue;
            }

            if (supply.Key.Length == 0)
            {
                missing += supplyRows.Count;
                warnings.Add(new ProcessingWarning(
                    "У " + supplyRows.Count + " строк не заполнена «Поставка» - файл согласования для них " +
                    "не найти, «" + PriceSchema.Prices.ApprovedPrice + "» остались как были.",
                    PriceSchema.PricesSheet,
                    sheetName,
                    new CellRef(firstDataRow + supplyRows[0].Index, priceColumn).ToString()));
                continue;
            }

            var needed = supplyRows.Select(row => row.Barcode).ToHashSet(StringComparer.Ordinal);
            var prices = new Dictionary<string, double>(StringComparer.Ordinal);
            var order = ApprovedPriceFiles.Order(supply.Key, files);
            var exactUsed = new List<string>();
            var nearestUsed = new List<string>();

            var sequence = order.Exact.Select(f => (f.Path, IsExact: true))
                .Concat(order.Nearest.Select(f => (f.Path, IsExact: false)))
                .Where(item => chosenFile is null || !SamePath(item.Path, chosenFile));
            if (chosenFile is not null)
            {
                sequence = sequence.Prepend((chosenFile, true));
            }

            foreach (var (path, isExact) in sequence)
            {
                if (needed.All(prices.ContainsKey))
                {
                    break;
                }

                var table = ReadTable(application, path);
                if (table is null)
                {
                    continue;
                }

                var added = 0;
                foreach (var barcode in needed)
                {
                    if (!prices.ContainsKey(barcode) && table.TryGetValue(barcode, out var price))
                    {
                        prices[barcode] = price;
                        added++;
                    }
                }

                if (isExact)
                {
                    exactUsed.Add(Path.GetFileName(path));
                }
                else if (added > 0)
                {
                    nearestUsed.Add(Path.GetFileName(path) + " - " + added + " шт.");
                }
            }

            foreach (var row in supplyRows)
            {
                if (prices.TryGetValue(row.Barcode, out var price))
                {
                    found[row.Index] = price;
                }
            }

            var notFound = supplyRows.Where(row => !prices.ContainsKey(row.Barcode)).ToList();
            missing += notFound.Count;

            _logger.Information(
                "Поставка " + supply.Key + ": цен найдено " + (supplyRows.Count - notFound.Count) + " из " +
                supplyRows.Count + ". Файл поставки: " + (exactUsed.Count == 0 ? "нет" : string.Join(", ", exactUsed)) +
                (nearestUsed.Count == 0 ? "" : "; из ближайших: " + string.Join(", ", nearestUsed)) + ".");

            var where = notFound.Count > 0
                ? new CellRef(firstDataRow + notFound[0].Index, priceColumn).ToString()
                : new CellRef(firstDataRow + supplyRows[0].Index, priceColumn).ToString();

            if (order.Exact.Count == 0 && chosenFile is null)
            {
                warnings.Add(new ProcessingWarning(
                    "Файл согласования цен для поставки " + supply.Key + " не найден в папке " + folder + "." +
                    (nearestUsed.Count > 0
                        ? " Цены взяты из ближайших по номеру поставок: " + string.Join(", ", nearestUsed) + " - проверьте."
                        : string.Empty),
                    PriceSchema.PricesSheet,
                    sheetName,
                    where));
            }
            else if (nearestUsed.Count > 0)
            {
                warnings.Add(new ProcessingWarning(
                    "В файле согласования поставки " + supply.Key + " нашлись не все штрихкоды, остальные цены " +
                    "взяты из ближайших по номеру поставок: " + string.Join(", ", nearestUsed) + " - проверьте.",
                    PriceSchema.PricesSheet,
                    sheetName,
                    where));
            }

            if (notFound.Count > 0)
            {
                warnings.Add(new ProcessingWarning(
                    "Согласованная цена не найдена для " + notFound.Count + " штрихкод(ов) поставки " + supply.Key +
                    ": " + string.Join(", ", notFound.Take(5).Select(row => row.Barcode)) +
                    (notFound.Count > 5 ? " и другие" : string.Empty) +
                    ". В этих строках «" + PriceSchema.Prices.ApprovedPrice + "» остались как были.",
                    PriceSchema.PricesSheet,
                    sheetName,
                    where));
            }
        }

        Write(pricesSheet, firstDataRow, priceColumn, found);
        _logger.Information("«Согласованные цены» проставлены: " + found.Count + ", не найдено: " + missing + ".");
        return new ApprovedPricesOutcome(found.Count, missing, warnings);
    }

    /// <summary>Файлы Excel в папке и её подпапках. Недоступные подпапки пропускаются.</summary>
    private IReadOnlyList<ApprovedPriceFile> ListFiles(string folder)
    {
        var options = new EnumerationOptions
        {
            RecurseSubdirectories = true,
            MaxRecursionDepth = FolderDepth,
            IgnoreInaccessible = true,
            MatchCasing = MatchCasing.CaseInsensitive,
        };

        try
        {
            return Directory.EnumerateFiles(folder, "*.xls*", options)
                .Where(ApprovedPriceFiles.IsWorkbook)
                .Select(path => new ApprovedPriceFile(path, File.GetLastWriteTime(path)))
                .ToList();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _logger.Warning("Не удалось просмотреть папку согласования цен " + folder + ".", ex);
            return Array.Empty<ApprovedPriceFile>();
        }
    }

    /// <summary>
    /// Цены файла по штрихкоду со всех листов, где нашлась таблица «ШК» + цена.
    /// null - файл не открылся или таблицы в нём нет.
    /// </summary>
    private IReadOnlyDictionary<string, double>? ReadTable(object applicationObject, string path)
    {
        if (_tables.TryGetValue(path, out var cached))
        {
            return cached;
        }

        dynamic application = applicationObject;
        using var scope = new ComScope();
        dynamic? book = null;
        int? security = null;
        Dictionary<string, double>? result = null;

        try
        {
            security = (int)application.AutomationSecurity;
            application.AutomationSecurity = AutomationSecurityForceDisable;

            dynamic workbooks = scope.Track(application.Workbooks);

            // Только чтение, без обновления связей и без вопросов: пароль пустой - защищённый
            // паролем файл не откроется, а не остановит обработку окном ввода пароля;
            // Notify = false - занятый другим человеком файл открывается копией, без ожидания.
            book = scope.Track(workbooks.Open(
                path, 0, true, Type.Missing, string.Empty, Type.Missing, true,
                Type.Missing, Type.Missing, false, false, Type.Missing, false));

            dynamic sheets = scope.Track(book.Worksheets);
            int count = sheets.Count;
            for (var index = 1; index <= count; index++)
            {
                dynamic sheet = scope.Track(sheets[index]);
                var bounds = ExcelSheetOperations.GetUsedBounds((object)sheet);
                if (bounds.RowCount < 2)
                {
                    continue;
                }

                var top = ExcelSheetOperations.ReadBlock(
                    (object)sheet,
                    bounds.FirstRow,
                    Math.Min(bounds.FirstRow + ApprovedPriceSheetReader.ScanRows - 1, bounds.LastRow),
                    bounds.FirstColumn,
                    bounds.LastColumn,
                    withFormulas: false);

                var table = ApprovedPriceSheetReader.FindTable(top);
                if (table is null || table.HeaderRow >= bounds.LastRow)
                {
                    continue;
                }

                var barcodes = ExcelSheetOperations.ReadBlock(
                    (object)sheet, table.HeaderRow + 1, bounds.LastRow, table.BarcodeColumn, table.BarcodeColumn, false);
                var prices = ExcelSheetOperations.ReadBlock(
                    (object)sheet, table.HeaderRow + 1, bounds.LastRow, table.PriceColumn, table.PriceColumn, false);

                result ??= new Dictionary<string, double>(StringComparer.Ordinal);
                var added = ApprovedPriceSheetReader.Read(barcodes, table.BarcodeColumn, prices, table.PriceColumn, result);
                _logger.Information(
                    "Файл согласования " + Path.GetFileName(path) + ", лист «" + (string)sheet.Name + "»: цена из «" +
                    table.PriceHeader + "», штрихкодов " + added + ".");
            }

            if (result is null)
            {
                _logger.Warning("В файле согласования " + Path.GetFileName(path) + " нет таблицы с колонками «ШК» и ценой.");
            }
        }
        catch (COMException ex)
        {
            _logger.Warning("Не удалось открыть файл согласования " + path + ".", ex);
            result = null;
        }
        finally
        {
            if (book is not null)
            {
                try
                {
                    book.Close(false);
                }
                catch (COMException ex)
                {
                    _logger.Warning("Не удалось закрыть файл согласования " + path + ".", ex);
                }
            }

            if (security is not null)
            {
                try
                {
                    application.AutomationSecurity = security.Value;
                }
                catch (COMException)
                {
                    // Настройка вернётся сама, когда закроется этот экземпляр Excel.
                }
            }
        }

        _tables[path] = result;
        return result;
    }

    /// <summary>
    /// Пишет найденные цены кусками подряд идущих строк. Строки, для которых цена не нашлась,
    /// не трогаются: в них может стоять то, что человек протянул руками.
    /// </summary>
    private static bool SamePath(string left, string right) =>
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);

    private static void Write(object sheet, int firstDataRow, int column, IReadOnlyDictionary<int, double> prices)
    {
        var indexes = prices.Keys.OrderBy(index => index).ToList();
        var start = 0;
        while (start < indexes.Count)
        {
            var end = start;
            while (end + 1 < indexes.Count && indexes[end + 1] == indexes[end] + 1)
            {
                end++;
            }

            var values = new List<object?>();
            for (var i = start; i <= end; i++)
            {
                values.Add(prices[indexes[i]]);
            }

            ExcelSheetOperations.SetColumnValues(sheet, firstDataRow + indexes[start], column, values);
            start = end + 1;
        }
    }
}
