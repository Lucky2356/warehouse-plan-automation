using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>Что сделала расстановка мест хранения.</summary>
internal sealed record RestockStorageOutcome(
    int Placed,
    int NotCollected,
    int Loaders,
    int LoaderQuantity,
    int ToApprove);

/// <summary>
/// Подтоварка, вторая часть инструкции: листы мест хранения, «Место хранения» на «на загрузку»,
/// «Не собрано», «Отдельно», листы «из‹место›» и загрузочники.
///
/// Данные мест хранения аналитик обновляет сама: «по адресам и таре», «Р» и «МПП». Программа
/// раскладывает их по листам «МП», «А», «В», «СЗП», считает, откуда брать каждую строку,
/// и собирает загрузочники. То, что требует решения человека - объединить маленькие заказы,
/// передать ли на склад столько, - остаётся за ним и попадает в замечания.
/// </summary>
internal sealed class ExcelRestockStorageStage
{
    private const int ChunkRows = 50000;

    /// <summary>Не хватило - ячейка «Места хранения» красная (RGB 255 124 128, в порядке BGR).</summary>
    private const int NotCollectedFill = 0x807CFF;

    /// <summary>Собрано с нескольких мест - текст красный (RGB 255 0 0).</summary>
    private const int SplitFont = 0x0000FF;

    /// <summary>Другой код, на котором хватает, - строка голубая (RGB 221 235 247).</summary>
    private const int ReplacedCodeFill = 0xF7EBDD;

    /// <summary>Количество разнесено по кодам - строки зелёные (RGB 226 239 218).</summary>
    private const int SplitCodeFill = 0xDAEFE2;

    private static readonly Regex LoaderSheetName = new(@"^З(МП|МПП|А|СЗП|В)(-.*|\s\d+)?$", RegexOptions.CultureInvariant);

    private readonly IAppLogger _logger;
    private readonly Func<DateTime> _now;

    public ExcelRestockStorageStage(IAppLogger logger, Func<DateTime> now)
    {
        _logger = logger;
        _now = now;
    }

    private sealed record Table(object Sheet, string Name, HeaderMap Headers, int Acr, int FirstRow, int LastRow, int LastColumn);

    private sealed record LoadRow(
        int Index,
        int ExcelRow,
        string AcrKey,
        double Quantity,
        IReadOnlyList<object?> PickValues,
        string Comment,
        string Sector,
        string Priority,
        string Note,
        object? ClientCode,
        object? Price);

    /// <summary>Остатки места: по АЦР - коды в порядке листа.</summary>
    private sealed class PlaceStock
    {
        public PlaceStock(string sheet, int acrColumn, int quantityColumn, int codeColumn)
        {
            Sheet = sheet;
            AcrColumn = acrColumn;
            QuantityColumn = quantityColumn;
            CodeColumn = codeColumn;
        }

        public string Sheet { get; }

        public int AcrColumn { get; }

        public int QuantityColumn { get; }

        public int CodeColumn { get; }

        public Dictionary<string, List<StockCode>> Codes { get; } = new(StringComparer.Ordinal);

        public void Add(string acrKey, string code, double quantity, string supplyNumber)
        {
            if (acrKey.Length == 0)
            {
                return;
            }

            if (!Codes.TryGetValue(acrKey, out var list))
            {
                list = new List<StockCode>();
                Codes[acrKey] = list;
            }

            var index = list.FindIndex(item => item.Code == code);
            if (index < 0)
            {
                list.Add(new StockCode(code, quantity, supplyNumber));
            }
            else
            {
                list[index] = list[index] with { Quantity = list[index].Quantity + quantity };
            }
        }

        public double Sum(string acrKey) => Codes.TryGetValue(acrKey, out var list) ? list.Sum(item => item.Quantity) : 0d;

        public IReadOnlyList<StockCode> Of(string acrKey) =>
            Codes.TryGetValue(acrKey, out var list) ? list : Array.Empty<StockCode>();
    }

    public RestockStorageOutcome? Run(
        object applicationObject,
        object workbookObject,
        object loadSheet,
        object anchorSheet,
        ComScope scope,
        Action<string, int> report,
        CancellationToken cancellationToken,
        List<ProcessingWarning> warnings)
    {
        var addressSheet = Lookup(workbookObject, RestockSchema.AddressSheet, scope, out var addressProblem);
        var reservesSheet = Lookup(workbookObject, RestockSchema.ReservesSheet, scope, out var reservesProblem);
        var suppliesSheet = Lookup(workbookObject, RestockSchema.SuppliesSheet, scope, out var suppliesProblem);

        var problems = new[] { addressProblem, reservesProblem, suppliesProblem }.Where(p => p is not null).ToList();
        if (problems.Count > 0)
        {
            warnings.Add(new ProcessingWarning(
                "Места хранения и загрузочники не собраны: " + string.Join("; ", problems) +
                ". «Заметка» и лист согласования готовы."));
            _logger.Warning("Расстановка мест хранения пропущена: " + string.Join("; ", problems));
            return null;
        }

        var listSeparator = ExcelSheetOperations.GetListSeparator(applicationObject);
        var decimalSeparator = DecimalSeparator(applicationObject);

        RemoveGeneratedSheets(workbookObject, scope);

        cancellationToken.ThrowIfCancellationRequested();
        report("Лист «" + RestockSchema.AddressSheet + "»: МП, А, В", 42);
        var stock = new Dictionary<StoragePlace, PlaceStock>();
        BuildAddressPlaces(applicationObject, workbookObject, addressSheet!, scope, decimalSeparator, stock, warnings);

        cancellationToken.ThrowIfCancellationRequested();
        report("Лист «" + RestockSchema.ReservesSheet + "»: резервы", 55);
        var reserves = ReadReserves(reservesSheet!, listSeparator, decimalSeparator, out var divisionGroups, out var reservesTable);

        cancellationToken.ThrowIfCancellationRequested();
        report("Лист «" + RestockSchema.SuppliesSheet + "»: МПП и СЗП", 60);
        BuildSupplyPlaces(workbookObject, suppliesSheet!, scope, listSeparator, decimalSeparator, stock, warnings);

        cancellationToken.ThrowIfCancellationRequested();
        report("Места хранения на «" + RestockSchema.LoadSheet + "»", 66);
        var load = ReadLoad(loadSheet, out var loadHeaders, out var loadFirst, out var loadLast);

        var allocations = new Dictionary<int, StorageAllocation>();
        foreach (var row in load.Where(row => row.AcrKey.Length > 0 && row.Quantity > 0d))
        {
            var sums = StoragePlaces.Priority.ToDictionary(place => place, place => stock[place].Sum(row.AcrKey));
            allocations[row.Index] = StorageAllocator.Allocate(
                row.Quantity,
                sums,
                reserves.TryGetValue(row.AcrKey, out var lines) ? lines : Array.Empty<ReserveLine>());
        }

        var loadName = ExcelSheetOperations.GetSheetName(loadSheet);
        WritePlaces(loadSheet, loadHeaders, loadFirst, loadLast, load, allocations, stock, reservesTable);
        AddUnknownReserveWarnings(load, allocations, loadHeaders, loadName, warnings);

        cancellationToken.ThrowIfCancellationRequested();
        report("Листы «Не собрано» и «Отдельно»", 72);
        var anchor = anchorSheet;
        var noteTitle = HeaderText(loadSheet, loadHeaders, RestockSchema.Load.Note);
        var notCollected = load.Where(row => allocations.TryGetValue(row.Index, out var a) && !a.Complete).ToList();
        if (notCollected.Count > 0)
        {
            anchor = WriteRowsSheet(
                workbookObject, anchor, loadSheet, loadHeaders, RestockSchema.NotCollectedSheet, notCollected, scope,
                new[] { "Не собрано", "Собрано", noteTitle, RestockSchema.Load.Place },
                row =>
                {
                    var allocation = allocations[row.Index];
                    return new object?[] { allocation.Missing, allocation.Collected, NullIfEmpty(row.Note), allocation.Text };
                });
        }

        var separate = load.Where(row => TextUtils.ContainsKey(row.Comment, RestockSchema.SeparateMarker)).ToList();
        if (separate.Count > 0)
        {
            anchor = WriteRowsSheet(
                workbookObject, anchor, loadSheet, loadHeaders, RestockSchema.SeparateSheet, separate, scope,
                new[] { noteTitle, RestockSchema.Load.Place },
                row => new object?[]
                {
                    NullIfEmpty(row.Note),
                    allocations.TryGetValue(row.Index, out var allocation) ? allocation.Text : null,
                });
        }

        cancellationToken.ThrowIfCancellationRequested();
        report("Листы «из‹место›»", 78);
        var byIndex = load.ToDictionary(row => row.Index);
        var pickLines = new List<PickLine>();
        foreach (var row in load)
        {
            if (!allocations.TryGetValue(row.Index, out var allocation))
            {
                continue;
            }

            foreach (var part in allocation.Parts)
            {
                pickLines.AddRange(PickListBuilder.Build(row.Index, part.Place, part.Quantity, stock[part.Place].Of(row.AcrKey)));
            }
        }

        var loaders = RestockLoaderBuilder.Build(pickLines, index => byIndex[index].Comment);

        // Загрузочники стоят сразу за «Не собрано» и «Отдельно»: их копируют
        // в «Распределительный логист», а листы «из‹место›» - рабочие, они дальше.
        foreach (var loader in loaders)
        {
            anchor = WriteLoader(workbookObject, anchor, loader, byIndex, divisionGroups, scope);
        }

        foreach (var place in StoragePlaces.Priority)
        {
            var lines = pickLines.Where(line => line.Place == place).ToList();
            if (lines.Count > 0)
            {
                anchor = WritePickSheet(
                    workbookObject, anchor, loadSheet, loadHeaders, place, lines, byIndex, allocations, stock[place], scope, warnings);
            }
        }

        // Сумма загрузочников равна «в подтоварку» за вычетом «Не собрано» по построению:
        // каждая часть «Места хранения» целиком раскладывается по кодам. Проверять здесь
        // нечего - нехватка на кодах уже пришла замечанием с листа «из‹место›».
        var loaderQuantity = loaders.Sum(loader => loader.Lines.Sum(line => line.Quantity));
        AddLoaderWarnings(loaders, byIndex, warnings);

        var toApprove = pickLines
            .Select(line => line.RowIndex)
            .Distinct()
            .Count(index => TextUtils.EqualsKey(byIndex[index].Note, TextUtils.NormalizeKey(RestockSchema.NoteApprove)));
        if (toApprove > 0)
        {
            warnings.Add(new ProcessingWarning(
                "В загрузочники попали строки с «Заметкой» «Согласовать»: " + toApprove +
                ". Грузите их только после согласования.",
                RestockSchema.ApprovalSheet,
                RestockSchema.ApprovalSheet));
        }

        _logger.Information(
            "Места хранения: расставлено " + allocations.Count + ", не собрано " + notCollected.Count +
            ", загрузочников " + loaders.Count + ", в них " + AllocationText(loaderQuantity) + " шт.");

        return new RestockStorageOutcome(
            allocations.Count,
            notCollected.Count,
            loaders.Count,
            (int)Math.Round(loaderQuantity),
            toApprove);
    }

    // ================= Листы мест хранения =================

    /// <summary>
    /// «по адресам и таре» → «МП», «А», «В». В первой колонке исходного листа - АЦР формулой
    /// «Артикул & Цвет & Размер», на новых листах - уже значением: по нему считаются СУММЕСЛИМН.
    /// </summary>
    private void BuildAddressPlaces(
        object applicationObject,
        object workbookObject,
        object sheet,
        ComScope scope,
        string decimalSeparator,
        Dictionary<StoragePlace, PlaceStock> stock,
        List<ProcessingWarning> warnings)
    {
        var table = OpenTable(sheet, RestockSchema.Stock.Specs);
        var headers = table.Headers;
        var article = headers[RestockSchema.Stock.Article];
        var color = headers[RestockSchema.Stock.Color];
        var size = headers[RestockSchema.Stock.Size];
        var code = headers[RestockSchema.Stock.Code];
        var quantity = headers[RestockSchema.Stock.Quantity];
        var type = headers[RestockSchema.Stock.StorageType];

        WriteAcrFormula(table, article, color, size);

        var places = new[] { StoragePlace.Marketplace, StoragePlace.Storage, StoragePlace.Returns };
        var rows = places.ToDictionary(place => place, _ => new List<object?[]>());
        foreach (var place in places)
        {
            stock[place] = new PlaceStock(StoragePlaces.Code(place), table.Acr, quantity, code);
        }

        var skipped = 0;
        ForEachRow(table, (grid, row) =>
        {
            var place = StoragePlaces.FromStorageType(grid.Text(row, type));
            if (place is null || !rows.ContainsKey(place.Value))
            {
                if (grid.NormalizedText(row, article).Length > 0)
                {
                    skipped++;
                }

                return;
            }

            var values = RowValues(grid, row, table.LastColumn);
            var acr = Acr(grid.Value(row, article), grid.Value(row, color), grid.Value(row, size), decimalSeparator);
            values[table.Acr - 1] = acr;
            rows[place.Value].Add(values);
            stock[place.Value].Add(
                TextUtils.NormalizeKey(acr),
                TextUtils.Normalize(grid.Text(row, code)),
                grid.Number(row, quantity) ?? 0d,
                string.Empty);
        });

        foreach (var place in places)
        {
            var name = StoragePlaces.Code(place);
            var created = ExcelSheetOperations.AddSheetBefore(workbookObject, sheet, name, scope);
            WriteTableSheet(table, created, rows[place]);
            _logger.Information("Лист «" + name + "»: строк " + rows[place].Count + ".");
        }

        _logger.Information(
            "«" + table.Name + "»: строк других типов хранения (образцы, брак, без маркировки) " + skipped + ".");
    }

    /// <summary>
    /// «Р»: остаются этот год, текущий месяц и два предыдущих, остальное удаляется. По оставшимся
    /// строкам резервы собираются по АЦР - вместе с комментарием заказа, по которому видно,
    /// с какого места хранения резерв.
    /// </summary>
    private Dictionary<string, IReadOnlyList<ReserveLine>> ReadReserves(
        object sheet,
        string listSeparator,
        string decimalSeparator,
        out Dictionary<string, string> divisionGroups,
        out Table table)
    {
        table = OpenTable(sheet, RestockSchema.Reserves.Specs);
        var headers = table.Headers;
        var date = headers[RestockSchema.Reserves.Date];
        var article = headers[RestockSchema.Reserves.Article];
        var color = headers[RestockSchema.Reserves.Color];
        var size = headers[RestockSchema.Reserves.Size];
        var quantity = headers[RestockSchema.Reserves.Quantity];
        var comment = headers[RestockSchema.Reserves.Comment];
        var top = ExcelSheetOperations.ReadBlock(table.Sheet, headers.HeaderRow, headers.HeaderRow, 1, table.LastColumn, withFormulas: false);
        var division = FindHeader(top, headers.HeaderRow, "подразделение");
        var group = FindHeader(top, headers.HeaderRow, "группа подразделения");
        var today = _now().Date;

        var toDelete = new List<int>();
        var result = new Dictionary<string, List<ReserveLine>>(StringComparer.Ordinal);
        var groups = new Dictionary<string, string>(StringComparer.Ordinal);

        ForEachRow(table, (grid, row) =>
        {
            var value = grid.Number(row, date);
            if (value is { } oa && oa > 0 && !ReservePeriod.Keep(DateTime.FromOADate(oa), today))
            {
                toDelete.Add(row);
                return;
            }

            var key = TextUtils.NormalizeKey(Acr(grid.Value(row, article), grid.Value(row, color), grid.Value(row, size), decimalSeparator));
            if (key.Length == 0)
            {
                return;
            }

            if (!result.TryGetValue(key, out var list))
            {
                list = new List<ReserveLine>();
                result[key] = list;
            }

            list.Add(new ReserveLine(grid.Number(row, quantity) ?? 0d, TextUtils.Normalize(grid.Text(row, comment))));
            if (division is { } divisionColumn && group is { } groupColumn)
            {
                groups.TryAdd(TextUtils.NormalizeKey(grid.Text(row, divisionColumn)), TextUtils.Normalize(grid.Text(row, groupColumn)));
            }
        });

        if (toDelete.Count > 0)
        {
            ExcelSheetOperations.DeleteRows(table.Sheet, toDelete, listSeparator);
            table = table with { LastRow = table.LastRow - toDelete.Count };
        }

        WriteAcrFormula(table, article, color, size);
        _logger.Information("«" + table.Name + "»: удалено строк прошлых месяцев " + toDelete.Count + ", АЦР с резервами " + result.Count + ".");

        divisionGroups = groups;
        return result.ToDictionary(pair => pair.Key, pair => (IReadOnlyList<ReserveLine>)pair.Value, StringComparer.Ordinal);
    }

    /// <summary>
    /// «МПП»: поставки на «С» переезжают на новый лист «СЗП» и удаляются с «МПП» - на нём
    /// остаются только «Л» и «М». Если «С» на «МПП» уже нет, а «СЗП» есть (книгу разбирают
    /// второй раз), остатки «СЗП» берутся с него.
    /// </summary>
    private void BuildSupplyPlaces(
        object workbookObject,
        object sheet,
        ComScope scope,
        string listSeparator,
        string decimalSeparator,
        Dictionary<StoragePlace, PlaceStock> stock,
        List<ProcessingWarning> warnings)
    {
        var table = OpenTable(sheet, RestockSchema.Supplies.Specs);
        var headers = table.Headers;
        var number = headers[RestockSchema.Supplies.Number];
        var article = headers[RestockSchema.Supplies.Article];
        var color = headers[RestockSchema.Supplies.Color];
        var size = headers[RestockSchema.Supplies.Size];
        var code = headers[RestockSchema.Supplies.Code];
        var deviation = headers[RestockSchema.Supplies.Deviation];

        stock[StoragePlace.Supplies] = new PlaceStock(table.Name, table.Acr, deviation, code);

        var network = new List<object?[]>();
        var networkRows = new List<int>();
        var other = 0;
        ForEachRow(table, (grid, row) =>
        {
            var supply = TextUtils.Normalize(grid.Text(row, number));
            var place = StoragePlaces.FromSupplyNumber(supply);
            if (place is null)
            {
                if (supply.Length > 0 || grid.NormalizedText(row, article).Length > 0)
                {
                    other++;
                }

                return;
            }

            var acr = Acr(grid.Value(row, article), grid.Value(row, color), grid.Value(row, size), decimalSeparator);
            if (place == StoragePlace.NetworkSupplies)
            {
                var values = RowValues(grid, row, table.LastColumn);
                values[table.Acr - 1] = acr;
                network.Add(values);
                networkRows.Add(row);
                return;
            }

            stock[StoragePlace.Supplies].Add(
                TextUtils.NormalizeKey(acr), TextUtils.Normalize(grid.Text(row, code)), grid.Number(row, deviation) ?? 0d, supply);
        });

        if (other > 0)
        {
            warnings.Add(new ProcessingWarning(
                "На листе «" + table.Name + "» " + other + " строк(и) с номером поставки не на «М», «Л» или «С»: " +
                "в остатки они не взяты.",
                table.Name,
                table.Name));
        }

        if (network.Count > 0)
        {
            ExcelSheetOperations.RemoveSheet(workbookObject, "СЗП", scope);
            var created = ExcelSheetOperations.AddSheet(workbookObject, sheet, "СЗП", scope);
            WriteTableSheet(table, created, network);
            ExcelSheetOperations.DeleteRows(sheet, networkRows, listSeparator);
            table = table with { LastRow = table.LastRow - networkRows.Count };

            var networkStock = new PlaceStock("СЗП", table.Acr, deviation, code);
            foreach (var values in network)
            {
                networkStock.Add(
                    TextUtils.NormalizeKey(TextUtils.CellToString(values[table.Acr - 1])),
                    TextUtils.Normalize(TextUtils.CellToString(values[code - 1])),
                    TextUtils.CellToDouble(values[deviation - 1]) ?? 0d,
                    TextUtils.Normalize(TextUtils.CellToString(values[number - 1])));
            }

            stock[StoragePlace.NetworkSupplies] = networkStock;
            _logger.Information("На лист «СЗП» перенесено строк: " + network.Count + ".");
        }
        else
        {
            stock[StoragePlace.NetworkSupplies] = ReadExistingNetwork(workbookObject, table, scope, decimalSeparator);
        }

        WriteAcrFormula(table, article, color, size);
    }

    private PlaceStock ReadExistingNetwork(object workbookObject, Table supplies, ComScope scope, string decimalSeparator)
    {
        var existing = ExistingSheet(workbookObject, "СЗП", scope);
        if (existing is null)
        {
            // Пустой «СЗП» всё равно нужен: на него ссылаются формулы «на загрузку».
            var created = ExcelSheetOperations.AddSheet(workbookObject, supplies.Sheet, "СЗП", scope);
            WriteTableSheet(supplies, created, Array.Empty<object?[]>());
            _logger.Information("Поставок на «С» нет ни на «МПП», ни на отдельном листе «СЗП»: «СЗП» создан пустым.");
            return new PlaceStock(
                "СЗП",
                supplies.Acr,
                supplies.Headers[RestockSchema.Supplies.Deviation],
                supplies.Headers[RestockSchema.Supplies.Code]);
        }

        var table = OpenTable(existing, RestockSchema.Supplies.Specs);
        var headers = table.Headers;
        var result = new PlaceStock(table.Name, table.Acr, headers[RestockSchema.Supplies.Deviation], headers[RestockSchema.Supplies.Code]);
        ForEachRow(table, (grid, row) => result.Add(
            TextUtils.NormalizeKey(Acr(
                grid.Value(row, headers[RestockSchema.Supplies.Article]),
                grid.Value(row, headers[RestockSchema.Supplies.Color]),
                grid.Value(row, headers[RestockSchema.Supplies.Size]),
                decimalSeparator)),
            TextUtils.Normalize(grid.Text(row, headers[RestockSchema.Supplies.Code])),
            grid.Number(row, headers[RestockSchema.Supplies.Deviation]) ?? 0d,
            TextUtils.Normalize(grid.Text(row, headers[RestockSchema.Supplies.Number]))));
        _logger.Information("Остатки «СЗП» взяты с уже собранного листа.");
        return result;
    }

    // ================= «на загрузку» =================

    private IReadOnlyList<LoadRow> ReadLoad(object sheet, out HeaderMap headers, out int firstRow, out int lastRow)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(sheet);
        var top = ExcelSheetOperations.ReadBlock(
            sheet, bounds.FirstRow, Math.Min(bounds.FirstRow + HeaderResolver.DefaultScanRows - 1, bounds.LastRow),
            bounds.FirstColumn, bounds.LastColumn, withFormulas: false);
        headers = RestockSchema.Load.ResolveHeaders(top);
        firstRow = headers.HeaderRow + 1;
        lastRow = ExcelSheetOperations.GetLastFilledRow(sheet, headers[RestockSchema.Load.Acr], bounds.LastRow);

        var marketplacePrice = FindHeader(top, headers.HeaderRow, "цена на мп");
        var lastColumn = Math.Max(bounds.LastColumn, headers.Columns.Values.Max());
        var grid = ExcelSheetOperations.ReadBlock(sheet, firstRow, Math.Max(lastRow, firstRow), 1, lastColumn, withFormulas: false);

        var result = new List<LoadRow>();
        for (var row = firstRow; row <= lastRow; row++)
        {
            var map = headers;
            var price = marketplacePrice is { } priceColumn && TextUtils.Normalize(grid.Text(row, priceColumn)).Length > 0
                ? grid.Value(row, priceColumn)
                : grid.Value(row, map[RestockSchema.Load.Price]);

            result.Add(new LoadRow(
                row - firstRow,
                row,
                TextUtils.NormalizeKey(grid.Text(row, map[RestockSchema.Load.Acr])),
                grid.Number(row, map[RestockSchema.Load.Quantity]) ?? 0d,
                RestockSchema.Load.PickColumns.Select(name => grid.Value(row, map[name])).ToList(),
                TextUtils.Normalize(grid.Text(row, map[RestockSchema.Load.Comment])),
                TextUtils.Normalize(grid.Text(row, map[RestockSchema.Load.Sector])),
                RestockLoaderBuilder.PriorityText(grid.Value(row, map[RestockSchema.Load.Priority])),
                TextUtils.Normalize(grid.Text(row, map[RestockSchema.Load.Note])),
                grid.Value(row, map[RestockSchema.Load.ClientCode]),
                price));
        }

        return result;
    }

    /// <summary>
    /// «Место хранения» и остатки мест по АЦР - формулами СУММЕСЛИМН, как в инструкции:
    /// «1МП», «3А», «5В», «2МПП», «4СЗП» (цифра - очередь, с которой место берётся) и «Р».
    /// Колонок нет - они дописываются справа. Не хватило - ячейка красная, собрано
    /// с нескольких мест - красный текст.
    /// </summary>
    private void WritePlaces(
        object sheet,
        HeaderMap headers,
        int firstRow,
        int lastRow,
        IReadOnlyList<LoadRow> load,
        IReadOnlyDictionary<int, StorageAllocation> allocations,
        IReadOnlyDictionary<StoragePlace, PlaceStock> stock,
        Table reserves)
    {
        if (lastRow < firstRow)
        {
            return;
        }

        var headerRow = headers.HeaderRow;
        var placeColumn = EnsureColumn(sheet, headers, RestockSchema.Load.Place, new[] { "место хранения", "мето хранения" });
        var columns = new List<(int Column, string Formula)>();

        var acrLetter = ExcelColumn.ToLetters(headers[RestockSchema.Load.Acr]);
        var ordered = new[]
        {
            StoragePlace.Marketplace, StoragePlace.Storage, StoragePlace.Returns, StoragePlace.Supplies, StoragePlace.NetworkSupplies,
        };

        foreach (var place in ordered)
        {
            var name = StoragePlaces.LoadColumn(place);
            var column = EnsureColumn(sheet, headers, name, new[] { name.ToLowerInvariant(), StoragePlaces.Code(place).ToLowerInvariant() });
            var placeStock = stock[place];
            columns.Add((column, SumIfs(placeStock.Sheet, placeStock.QuantityColumn, placeStock.AcrColumn, acrLetter, firstRow)));
        }

        var reservesColumn = EnsureColumn(sheet, headers, RestockSchema.Load.Reserves, new[] { "р" });
        columns.Add((reservesColumn, SumIfs(
            reserves.Name, reserves.Headers[RestockSchema.Reserves.Quantity], reserves.Acr, acrLetter, firstRow)));

        foreach (var (column, formula) in columns)
        {
            SetRangeFormula(sheet, firstRow, lastRow, column, formula);
        }

        var texts = load.Select(row => allocations.TryGetValue(row.Index, out var allocation) ? (object?)allocation.Text : null).ToList();

        using (var scope = new ComScope())
        {
            dynamic target = scope.Track(((dynamic)sheet).Range[Reference(firstRow, placeColumn, lastRow, placeColumn)]);
            dynamic interior = scope.Track(target.Interior);
            interior.ColorIndex = ExcelConstants.XlColorIndexNone;
            dynamic font = scope.Track(target.Font);
            font.ColorIndex = ExcelConstants.XlColorIndexAutomatic;
        }

        WriteColumn(sheet, firstRow, placeColumn, texts);

        foreach (var row in load)
        {
            if (!allocations.TryGetValue(row.Index, out var allocation))
            {
                continue;
            }

            if (!allocation.Complete)
            {
                ExcelSheetOperations.SetFill(sheet, row.ExcelRow, placeColumn, NotCollectedFill);
            }
            else if (!allocation.SinglePlace)
            {
                ExcelSheetOperations.SetTextColor(sheet, row.ExcelRow, placeColumn, 0, allocation.Text.Length, SplitFont);
            }
        }

        _logger.Information(
            "«Место хранения» на строке заголовков " + headerRow + ", колонка " + ExcelColumn.ToLetters(placeColumn) + ".");
    }

    private void AddUnknownReserveWarnings(
        IReadOnlyList<LoadRow> load,
        IReadOnlyDictionary<int, StorageAllocation> allocations,
        HeaderMap headers,
        string sheetName,
        List<ProcessingWarning> warnings)
    {
        foreach (var row in load)
        {
            if (!allocations.TryGetValue(row.Index, out var allocation) || allocation.UnknownReserves.Count == 0)
            {
                continue;
            }

            var reserve = allocation.UnknownReserves[0];
            warnings.Add(new ProcessingWarning(
                "АЦР " + TextUtils.CellToString(row.PickValues[3]) + ": по комментарию резерва «" + Shorten(reserve.Comment) +
                "» (" + AllocationText(allocation.UnknownReserves.Sum(r => r.Quantity)) +
                " шт.) не понять, с какого он места хранения, - из остатков он не вычтен. Проверьте «" +
                allocation.Text + "».",
                RestockSchema.LoadSheet + ", строка " + row.ExcelRow,
                sheetName,
                ExcelColumn.ToLetters(headers[RestockSchema.Load.Acr]) + row.ExcelRow));
        }
    }

    // ================= Новые листы =================

    /// <summary>«Не собрано» и «Отдельно»: колонки строки до «Приоритета» и свои колонки справа.</summary>
    private object WriteRowsSheet(
        object workbookObject,
        object anchor,
        object loadSheet,
        HeaderMap loadHeaders,
        string name,
        IReadOnlyList<LoadRow> rows,
        ComScope scope,
        IReadOnlyList<string> extraHeaders,
        Func<LoadRow, object?[]> extraValues)
    {
        var sheet = ExcelSheetOperations.AddSheet(workbookObject, anchor, name, scope);
        CopyPickHeaders(loadSheet, loadHeaders, sheet);

        var first = RestockSchema.Load.PickColumns.Count + 1;
        for (var i = 0; i < extraHeaders.Count; i++)
        {
            SetHeader(sheet, first + i, extraHeaders[i]);
        }

        for (var c = 0; c < RestockSchema.Load.PickColumns.Count; c++)
        {
            var index = c;
            WriteColumn(sheet, 2, c + 1, rows.Select(row => row.PickValues[index]).ToList());
        }

        var extras = rows.Select(extraValues).ToList();
        for (var i = 0; i < extraHeaders.Count; i++)
        {
            var index = i;
            WriteColumn(sheet, 2, first + index, extras.Select(values => values[index]).ToList());
        }

        FinishSheet(sheet, rows.Count + 1, first + extraHeaders.Count - 1);
        _logger.Information("Лист «" + name + "»: строк " + rows.Count + ".");
        return sheet;
    }

    /// <summary>
    /// Лист «из‹место›»: строка на каждый код. «Код с адресов» - код места хранения,
    /// «Количество на этом коде» и «Остаток за вычетом количеств в загрузку» - формулами.
    /// </summary>
    private object WritePickSheet(
        object workbookObject,
        object anchor,
        object loadSheet,
        HeaderMap loadHeaders,
        StoragePlace place,
        IReadOnlyList<PickLine> lines,
        IReadOnlyDictionary<int, LoadRow> rows,
        IReadOnlyDictionary<int, StorageAllocation> allocations,
        PlaceStock stock,
        ComScope scope,
        List<ProcessingWarning> warnings)
    {
        var name = StoragePlaces.PickSheet(place);
        var sheet = ExcelSheetOperations.AddSheet(workbookObject, anchor, name, scope);
        CopyPickHeaders(loadSheet, loadHeaders, sheet);

        var pick = RestockSchema.Load.PickColumns;
        var codeIndex = IndexOf(pick, RestockSchema.Load.Code);
        var quantityIndex = IndexOf(pick, RestockSchema.Load.Quantity);

        var codeColumn = pick.Count + 1;
        var onCodeColumn = pick.Count + 2;
        var restColumn = pick.Count + 3;
        var supplyColumn = StoragePlaces.IsSupply(place) ? pick.Count + 4 : (int?)null;
        var placeColumn = (supplyColumn ?? restColumn) + 1;

        SetHeader(sheet, codeColumn, "Код с адресов " + StoragePlaces.Code(place));
        SetHeader(sheet, onCodeColumn, "Количество на этом коде");
        SetHeader(sheet, restColumn, "Остаток за вычетом количеств в загрузку");
        if (supplyColumn is { } supplyHeader)
        {
            SetHeader(sheet, supplyHeader, "Поставка");
        }

        SetHeader(sheet, placeColumn, RestockSchema.Load.Place);

        for (var c = 0; c < pick.Count; c++)
        {
            var index = c;
            WriteColumn(sheet, 2, c + 1, lines.Select(line => index == codeIndex
                    ? CodeValue(line.Code)
                    : index == quantityIndex
                        ? line.Quantity
                        : rows[line.RowIndex].PickValues[index])
                .ToList());
        }

        WriteColumn(sheet, 2, codeColumn, lines.Select(line => CodeValue(line.Code)).ToList());
        if (supplyColumn is { } supply)
        {
            WriteColumn(sheet, 2, supply, lines.Select(line => NullIfEmpty(line.SupplyNumber)).ToList());
        }

        WriteColumn(sheet, 2, placeColumn, lines.Select(line => (object?)allocations[line.RowIndex].Text).ToList());

        var last = lines.Count + 1;
        var codeLetter = ExcelColumn.ToLetters(codeColumn);
        SetRangeFormula(
            sheet, 2, last, onCodeColumn,
            "=SUMIFS(" + Column(stock.Sheet, stock.QuantityColumn) + "," + Column(stock.Sheet, stock.CodeColumn) + "," + codeLetter + "2)");
        SetRangeFormula(
            sheet, 2, last, restColumn,
            "=" + ExcelColumn.ToLetters(onCodeColumn) + "2-" + ExcelColumn.ToLetters(quantityIndex + 1) + "2");

        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines[i];
            var excelRow = i + 2;
            var fill = line.Mark switch
            {
                PickMark.ReplacedCode => ReplacedCodeFill,
                PickMark.SplitCode => SplitCodeFill,
                _ => (int?)null,
            };

            if (fill is { } color)
            {
                using var fillScope = new ComScope();
                dynamic range = fillScope.Track(((dynamic)sheet).Range[Reference(excelRow, 1, excelRow, placeColumn)]);
                dynamic interior = fillScope.Track(range.Interior);
                interior.Color = color;
            }

            if (line.Shortage > 0d)
            {
                warnings.Add(new ProcessingWarning(
                    "АЦР " + TextUtils.CellToString(rows[line.RowIndex].PickValues[3]) + ": на кодах «" + StoragePlaces.Code(place) +
                    "» не хватает " + AllocationText(line.Shortage) + " шт. - в «Остатке за вычетом» минус, проверьте по АЦР вручную.",
                    name + ", строка " + excelRow,
                    name,
                    ExcelColumn.ToLetters(restColumn) + excelRow));
            }
        }

        FinishSheet(sheet, last, placeColumn);
        _logger.Information(
            "Лист «" + name + "»: строк " + lines.Count + ", другой код " + lines.Count(l => l.Mark == PickMark.ReplacedCode) +
            ", разнесено по кодам " + lines.Count(l => l.Mark == PickMark.SplitCode) + ".");
        return sheet;
    }

    /// <summary>Загрузочник: «Код», «Подразделение», «Количество», «Цена», «Комментарий», «Номер заказа».</summary>
    private object WriteLoader(
        object workbookObject,
        object anchor,
        RestockLoader loader,
        IReadOnlyDictionary<int, LoadRow> rows,
        IReadOnlyDictionary<string, string> divisionGroups,
        ComScope scope)
    {
        var sheet = ExcelSheetOperations.AddSheet(workbookObject, anchor, loader.SheetName, scope);
        var first = rows[loader.Lines[0].RowIndex];
        var clientCode = TextUtils.CellToString(first.ClientCode);
        divisionGroups.TryGetValue(TextUtils.NormalizeKey(clientCode), out var group);

        var comment = RestockLoaderBuilder.CommentText(
            RestockLoaderBuilder.MarketplaceName(clientCode, group),
            loader.Lines.Select(line => rows[line.RowIndex].Sector),
            loader.Place,
            loader.Lines.Select(line => line.SupplyNumber),
            loader.Lines.Select(line => rows[line.RowIndex].Priority));

        var headers = new[] { "Код", "Подразделение", "Количество", "Цена", "Комментарий", "Номер заказа" };
        for (var c = 0; c < headers.Length; c++)
        {
            SetHeader(sheet, c + 1, headers[c]);
        }

        WriteColumn(sheet, 2, 1, loader.Lines.Select(line => CodeValue(line.Code)).ToList());
        WriteColumn(sheet, 2, 2, loader.Lines.Select(line => rows[line.RowIndex].ClientCode).ToList());
        WriteColumn(sheet, 2, 3, loader.Lines.Select(line => (object?)line.Quantity).ToList());
        WriteColumn(sheet, 2, 4, loader.Lines.Select(line => rows[line.RowIndex].Price).ToList());
        WriteColumn(sheet, 2, 5, loader.Lines.Select(_ => (object?)comment).ToList());
        WriteColumn(sheet, 2, 6, loader.Lines.Select(_ => (object?)1d).ToList());

        ExcelSheetOperations.AutoFitColumns(sheet, 1, headers.Length);
        _logger.Information("Загрузочник «" + loader.SheetName + "»: строк " + loader.Lines.Count + ". " + comment);
        return sheet;
    }

    private static void AddLoaderWarnings(
        IReadOnlyList<RestockLoader> loaders,
        IReadOnlyDictionary<int, LoadRow> rows,
        List<ProcessingWarning> warnings)
    {
        foreach (var loader in loaders)
        {
            var quantity = loader.Lines.Sum(line => line.Quantity);
            var sectors = string.Join(", ", loader.Lines
                .Select(line => rows[line.RowIndex].Sector)
                .Where(sector => sector.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase));

            warnings.Add(new ProcessingWarning(
                "Загрузочник «" + loader.SheetName + "»: строк " + loader.Lines.Count + ", " + AllocationText(quantity) + " шт." +
                (sectors.Length > 0 ? " (" + sectors + ")" : string.Empty) +
                ". Проверьте, можно ли столько передать на склад, не объединить ли с другим заказом, и текст комментария.",
                loader.SheetName,
                loader.SheetName,
                "E2"));
        }
    }

    // ================= Помощники =================

    private static object? Lookup(object workbookObject, string name, ComScope scope, out string? problem)
    {
        var lookup = ExcelSheetOperations.FindSheet(workbookObject, name, scope);
        problem = lookup.Sheet is not null
            ? null
            : lookup.Candidates.Count == 0
                ? "нет листа «" + name + "»"
                : "несколько листов «" + name + "»: " + string.Join(", ", lookup.Candidates);
        return lookup.Sheet;
    }

    private static object? ExistingSheet(object workbookObject, string name, ComScope scope)
    {
        dynamic workbook = workbookObject;
        dynamic sheets = scope.Track(workbook.Worksheets);
        int count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            dynamic sheet = sheets[index];
            if (string.Equals((string)sheet.Name, name, StringComparison.Ordinal))
            {
                scope.Track(sheet);
                return (object)sheet;
            }

            ComUtils.Release(sheet);
        }

        return null;
    }

    /// <summary>
    /// Прошлые листы этого шага удаляются заранее: «Не собрано», «Отдельно», «из‹место›»,
    /// загрузочники и «МП», «А», «В». «СЗП» удаляется, только когда пересоздаётся:
    /// при повторном разборе поставок на «С» на «МПП» уже нет.
    /// </summary>
    private void RemoveGeneratedSheets(object workbookObject, ComScope scope)
    {
        var names = new List<string> { RestockSchema.NotCollectedSheet, RestockSchema.SeparateSheet, "МП", "А", "В" };
        names.AddRange(StoragePlaces.Priority.Select(StoragePlaces.PickSheet));

        dynamic workbook = workbookObject;
        dynamic sheets = scope.Track(workbook.Worksheets);
        int count = sheets.Count;
        var removed = new List<string>();

        for (var index = count; index >= 1; index--)
        {
            dynamic sheet = sheets[index];
            string sheetName = sheet.Name;
            if (names.Contains(sheetName, StringComparer.Ordinal) || LoaderSheetName.IsMatch(sheetName))
            {
                sheet.Delete();
                removed.Add(sheetName);
            }

            ComUtils.Release(sheet);
        }

        if (removed.Count > 0)
        {
            _logger.Information("Удалены листы прошлого разбора: " + string.Join(", ", removed) + ".");
        }
    }

    /// <summary>
    /// Таблица листа данных. Колонки «АЦР» нет - она вставляется первой: так её и заводит
    /// инструкция, а СУММЕСЛИМН на «на загрузку» считают по ней.
    /// </summary>
    private Table OpenTable(object sheet, IReadOnlyList<ColumnSpec> specs)
    {
        ExcelSheetOperations.ShowAllRows(sheet);
        var name = ExcelSheetOperations.GetSheetName(sheet);
        var (headers, grid, bounds) = ReadHeaders(sheet, name, specs);

        var acr = FindHeader(grid, headers.HeaderRow, "ацр");
        if (acr is null)
        {
            using (var scope = new ComScope())
            {
                dynamic column = scope.Track(((dynamic)sheet).Range["A:A"]);
                dynamic entire = scope.Track(column.EntireColumn);
                entire.Insert();
            }

            ExcelSheetOperations.SetValue(sheet, headers.HeaderRow, 1, RestockSchema.Stock.Acr);
            _logger.Information("На лист «" + name + "» добавлена колонка «АЦР».");
            (headers, grid, bounds) = ReadHeaders(sheet, name, specs);
            acr = 1;
        }

        var lastColumn = Math.Max(headers.Columns.Values.Max(), acr.Value);
        for (var column = grid.LastColumn; column > lastColumn; column--)
        {
            if (grid.NormalizedText(headers.HeaderRow, column).Length > 0)
            {
                lastColumn = column;
                break;
            }
        }

        var lastRow = headers.Columns.Values
            .Select(column => ExcelSheetOperations.GetLastFilledRow(sheet, column, bounds.LastRow))
            .Max();

        return new Table(sheet, name, headers, acr.Value, headers.HeaderRow + 1, lastRow, lastColumn);
    }

    private static (HeaderMap Headers, SheetGrid Grid, SheetBounds Bounds) ReadHeaders(
        object sheet, string name, IReadOnlyList<ColumnSpec> specs)
    {
        var bounds = ExcelSheetOperations.GetUsedBounds(sheet);
        var grid = ExcelSheetOperations.ReadBlock(
            sheet, bounds.FirstRow, Math.Min(bounds.FirstRow + HeaderResolver.DefaultScanRows - 1, bounds.LastRow),
            1, Math.Max(bounds.LastColumn, 1), withFormulas: false);
        return (HeaderResolver.Resolve(grid, name, specs), grid, bounds);
    }

    private static int? FindHeader(SheetGrid grid, int headerRow, string key)
    {
        for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
        {
            if (TextUtils.EqualsKey(grid.Text(headerRow, column), key))
            {
                return column;
            }
        }

        return null;
    }

    private static void ForEachRow(Table table, Action<SheetGrid, int> action)
    {
        for (var start = table.FirstRow; start <= table.LastRow; start += ChunkRows)
        {
            var end = Math.Min(start + ChunkRows - 1, table.LastRow);
            var grid = ExcelSheetOperations.ReadBlock(table.Sheet, start, end, 1, table.LastColumn, withFormulas: false);
            for (var row = start; row <= end; row++)
            {
                action(grid, row);
            }
        }
    }

    private static object?[] RowValues(SheetGrid grid, int row, int lastColumn)
    {
        var values = new object?[lastColumn];
        for (var column = 1; column <= lastColumn; column++)
        {
            values[column - 1] = grid.Value(row, column);
        }

        return values;
    }

    /// <summary>
    /// АЦР так, как его склеивает формула «Артикул & Цвет & Размер»: дробный размер Excel пишет
    /// с десятичным разделителем системы.
    /// </summary>
    private static string Acr(object? article, object? color, object? size, string decimalSeparator)
    {
        static string Part(object? value, string separator) => value is double number && number != Math.Floor(number)
            ? number.ToString("R", CultureInfo.InvariantCulture).Replace(".", separator, StringComparison.Ordinal)
            : TextUtils.CellToString(value);

        return Part(article, decimalSeparator) + Part(color, decimalSeparator) + Part(size, decimalSeparator);
    }

    private void WriteAcrFormula(Table table, int article, int color, int size)
    {
        if (table.LastRow < table.FirstRow)
        {
            return;
        }

        var row = table.FirstRow.ToString(CultureInfo.InvariantCulture);
        var formula = "=" + ExcelColumn.ToLetters(article) + row + "&" + ExcelColumn.ToLetters(color) + row + "&" +
                      ExcelColumn.ToLetters(size) + row;
        SetRangeFormula(table.Sheet, table.FirstRow, table.LastRow, table.Acr, formula);
    }

    /// <summary>Новый лист места хранения: шапка и ширины колонок исходного листа, строки значениями.</summary>
    private static void WriteTableSheet(Table source, object target, IReadOnlyList<object?[]> rows)
    {
        using (var scope = new ComScope())
        {
            dynamic from = scope.Track(((dynamic)source.Sheet).Range[Reference(source.Headers.HeaderRow, 1, source.Headers.HeaderRow, source.LastColumn)]);
            dynamic to = scope.Track(((dynamic)target).Range["A1"]);
            from.Copy(to);

            for (var column = 1; column <= source.LastColumn; column++)
            {
                dynamic sourceColumn = scope.Track(((dynamic)source.Sheet).Columns[column]);
                dynamic targetColumn = scope.Track(((dynamic)target).Columns[column]);
                targetColumn.ColumnWidth = sourceColumn.ColumnWidth;
            }
        }

        for (var column = 1; column <= source.LastColumn; column++)
        {
            var index = column - 1;
            WriteColumn(target, 2, column, rows.Select(values => values[index]).ToList());
        }

        FinishSheet(target, rows.Count + 1, source.LastColumn);
    }

    private static void CopyPickHeaders(object loadSheet, HeaderMap loadHeaders, object target)
    {
        using var scope = new ComScope();
        for (var i = 0; i < RestockSchema.Load.PickColumns.Count; i++)
        {
            var source = loadHeaders[RestockSchema.Load.PickColumns[i]];
            dynamic from = scope.Track(((dynamic)loadSheet).Cells[loadHeaders.HeaderRow, source]);
            dynamic to = scope.Track(((dynamic)target).Cells[1, i + 1]);
            from.Copy(to);

            dynamic sourceColumn = scope.Track(((dynamic)loadSheet).Columns[source]);
            dynamic targetColumn = scope.Track(((dynamic)target).Columns[i + 1]);
            targetColumn.ColumnWidth = sourceColumn.ColumnWidth;
        }
    }

    /// <summary>Заголовок колонки так, как он написан на листе: «Заметка» или «Заметки».</summary>
    private static string HeaderText(object sheet, HeaderMap headers, string name)
    {
        var text = TextUtils.Normalize(TextUtils.CellToString(
            ExcelSheetOperations.GetValue(sheet, headers.HeaderRow, headers[name])));
        return text.Length > 0 ? text : name;
    }

    private static void SetHeader(object sheet, int column, string text)
    {
        using var scope = new ComScope();
        dynamic cell = scope.Track(((dynamic)sheet).Cells[1, column]);
        cell.Value2 = text;
        dynamic font = scope.Track(cell.Font);
        font.Bold = true;
        cell.WrapText = true;
    }

    /// <summary>Автофильтр на шапку - по этим листам дальше фильтруются.</summary>
    private static void FinishSheet(object sheet, int lastRow, int lastColumn)
    {
        using var scope = new ComScope();
        dynamic range = scope.Track(((dynamic)sheet).Range[Reference(1, 1, Math.Max(lastRow, 2), lastColumn)]);
        range.AutoFilter();
    }

    /// <summary>
    /// Колонка «на загрузку» по названию; нет - дописывается справа от последней колонки
    /// с заголовком, в оформлении заголовка «Заметки».
    /// </summary>
    private static int EnsureColumn(object sheet, HeaderMap headers, string title, IReadOnlyList<string> keys) =>
        ExcelSheetOperations.EnsureHeaderColumn(
            sheet,
            headers.HeaderRow,
            headers.Columns.Values.Max(),
            title,
            keys,
            headers.TryGet(RestockSchema.Load.Note, out var note) ? note : 0);

    private static string SumIfs(string sheet, int sumColumn, int criteriaColumn, string acrLetter, int row) =>
        "=SUMIFS(" + Column(sheet, sumColumn) + "," + Column(sheet, criteriaColumn) + ",$" + acrLetter +
        row.ToString(CultureInfo.InvariantCulture) + ")";

    private static string Column(string sheet, int column)
    {
        var letters = ExcelColumn.ToLetters(column);
        return "'" + sheet.Replace("'", "''", StringComparison.Ordinal) + "'!$" + letters + ":$" + letters;
    }

    private static void SetRangeFormula(object sheet, int firstRow, int lastRow, int column, string formula)
    {
        if (lastRow < firstRow)
        {
            return;
        }

        using var scope = new ComScope();
        dynamic range = scope.Track(((dynamic)sheet).Range[Reference(firstRow, column, lastRow, column)]);
        range.Formula = formula;
    }

    /// <summary>
    /// Колонка значений. Текст, который Excel принял бы за число или дату («352-148», «56-58»),
    /// пишется в текстовом формате - иначе значение изменилось бы.
    /// </summary>
    private static void WriteColumn(object sheet, int firstRow, int column, IReadOnlyList<object?> values)
    {
        if (values.Count == 0)
        {
            return;
        }

        if (values.Any(value => value is string text && (CellError.LooksNumericToExcel(text) || text.StartsWith("=", StringComparison.Ordinal))))
        {
            ExcelSheetOperations.SetRangeAsText(sheet, firstRow, firstRow + values.Count - 1, column, column);
        }

        ExcelSheetOperations.SetColumnValues(sheet, firstRow, column, values);
    }

    /// <summary>Код товара числом, если он число: так он стоит на листах мест хранения.</summary>
    private static object? CodeValue(string code) =>
        code.Length > 0 && code.All(char.IsDigit) && code[0] != '0' && code.Length < 16
            ? double.Parse(code, CultureInfo.InvariantCulture)
            : NullIfEmpty(code);

    private static object? NullIfEmpty(string? text) => string.IsNullOrEmpty(text) ? null : text;

    private static int IndexOf(IReadOnlyList<string> list, string value)
    {
        for (var i = 0; i < list.Count; i++)
        {
            if (list[i] == value)
            {
                return i;
            }
        }

        return -1;
    }

    private static string AllocationText(double value) => StorageAllocation.Quantity(value);

    private static string Shorten(string text) => text.Length <= 60 ? text : text[..60] + "…";

    private static string DecimalSeparator(object applicationObject)
    {
        dynamic application = applicationObject;
        try
        {
            object? value = application.International(ExcelConstants.XlDecimalSeparator);
            return value as string is { Length: > 0 } separator ? separator : ",";
        }
        catch (COMException)
        {
            return ",";
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            return ",";
        }
    }

    private static string Reference(int firstRow, int firstColumn, int lastRow, int lastColumn) =>
        ExcelColumn.ToLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture) + ":" +
        ExcelColumn.ToLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture);
}
