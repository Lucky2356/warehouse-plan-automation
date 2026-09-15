using System.Text.RegularExpressions;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Номер поставки, разобранный на начало и порядковую часть: «431-329» - «431» и 329.</summary>
public readonly record struct SupplyKey(string Prefix, int Number);

/// <summary>Файл в папке согласования цен.</summary>
public sealed record ApprovedPriceFile(string Path, DateTime Modified);

/// <summary>
/// В каком порядке смотреть файлы согласования для поставки: сначала файлы с её номером,
/// потом ближайшие по номеру поставки с тем же началом.
/// </summary>
public sealed record ApprovedPriceFileOrder(
    IReadOnlyList<ApprovedPriceFile> Exact,
    IReadOnlyList<ApprovedPriceFile> Nearest);

/// <summary>
/// Поиск файла согласования цен по номеру поставки.
///
/// Вручную это так: в папке КМ «Согласование цен» → папка сезона ищется файл с номером
/// поставки. Если его нет или в нём нашлись не все штрихкоды - берётся ближайший файл
/// «по первым цифрам поставки»: у «431-329» это «431-328», «431-330» и так далее.
/// </summary>
public static class ApprovedPriceFiles
{
    /// <summary>Сколько ближайших файлов просматривается, если в файле поставки нашлось не всё.</summary>
    public const int NearestLimit = 10;

    private static readonly Regex KeyPattern = new(
        @"(?<![\p{L}\d])(?<prefix>\p{L}{0,3}\d+)\s?-\s?(?<number>\d{1,6})(?!\d)",
        RegexOptions.CultureInvariant);

    /// <summary>Номер поставки, если он похож на номер: «431-329», «C2518-084».</summary>
    public static SupplyKey? Parse(string? supply)
    {
        var match = KeyPattern.Match(TextUtils.Normalize(supply));
        return match.Success ? ToKey(match) : null;
    }

    /// <summary>Все номера поставок, которые встречаются в названии файла.</summary>
    public static IReadOnlyList<SupplyKey> KeysInFileName(string path)
    {
        var name = System.IO.Path.GetFileNameWithoutExtension(path);
        return KeyPattern.Matches(name).Select(ToKey).ToList();
    }

    public static ApprovedPriceFileOrder Order(string supply, IEnumerable<ApprovedPriceFile> files)
    {
        var key = Parse(supply);
        if (key is null)
        {
            // Номер не разобрался - остаётся искать его как есть в названии файла.
            var text = TextUtils.NormalizeKey(supply);
            var byText = text.Length == 0
                ? new List<ApprovedPriceFile>()
                : files
                    .Where(file => TextUtils.ContainsKey(System.IO.Path.GetFileNameWithoutExtension(file.Path), text))
                    .OrderByDescending(file => file.Modified)
                    .ToList();
            return new ApprovedPriceFileOrder(byText, Array.Empty<ApprovedPriceFile>());
        }

        var exact = new List<ApprovedPriceFile>();
        var nearest = new List<(ApprovedPriceFile File, int Distance)>();

        foreach (var file in files)
        {
            var keys = KeysInFileName(file.Path).Where(k => k.Prefix == key.Value.Prefix).ToList();
            if (keys.Count == 0)
            {
                continue;
            }

            if (keys.Any(k => k.Number == key.Value.Number))
            {
                exact.Add(file);
                continue;
            }

            nearest.Add((file, keys.Min(k => Math.Abs(k.Number - key.Value.Number))));
        }

        return new ApprovedPriceFileOrder(
            exact.OrderByDescending(file => file.Modified).ToList(),
            nearest
                .OrderBy(item => item.Distance)
                .ThenByDescending(item => item.File.Modified)
                .Take(NearestLimit)
                .Select(item => item.File)
                .ToList());
    }

    /// <summary>Файл Excel, а не временный файл открытой книги («~$...»).</summary>
    public static bool IsWorkbook(string path)
    {
        var name = System.IO.Path.GetFileName(path);
        return !name.StartsWith("~$", StringComparison.Ordinal) && WorkbookFileExtension(path);
    }

    private static bool WorkbookFileExtension(string path)
    {
        var extension = System.IO.Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsb", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xls", StringComparison.OrdinalIgnoreCase);
    }

    private static SupplyKey ToKey(Match match) =>
        new(SameLetters(match.Groups["prefix"].Value), int.Parse(match.Groups["number"].Value, System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>
    /// Буква в номере поставки набирается то латиницей, то кириллицей: «C2518» и «С2518»
    /// выглядят одинаково, поэтому и сравниваются одинаково.
    /// </summary>
    private static string SameLetters(string prefix)
    {
        const string latin = "ABCEHKMOPTXY";
        const string cyrillic = "АВСЕНКМОРТХУ";

        var upper = prefix.ToUpperInvariant().ToCharArray();
        for (var i = 0; i < upper.Length; i++)
        {
            var index = latin.IndexOf(upper[i]);
            if (index >= 0)
            {
                upper[i] = cyrillic[index];
            }
        }

        return new string(upper);
    }
}

/// <summary>Где в файле согласования лежит таблица: строка заголовков и две нужные колонки.</summary>
public sealed record ApprovedPriceTable(int HeaderRow, int BarcodeColumn, int PriceColumn, string PriceHeader);

/// <summary>
/// Таблица файла согласования цен. Нужны две колонки: «ШК» и цена. Ценой считается
/// «Розничная цена», а если её нет - «Предложение УТЗ»: в инструкции ВПР берёт диапазон
/// «с «ШК» до «Розничная цена», либо «Предложение УТЗ»».
/// </summary>
public static class ApprovedPriceSheetReader
{
    /// <summary>Сколько первых строк листа просматривается в поиске заголовков.</summary>
    public const int ScanRows = 30;

    private static readonly string[] BarcodeHeaders = { "шк", "штрихкод", "штрих-код", "штрих код" };

    private static readonly string[] PriceHeaders = { "розничная цена", "предложение утз" };

    public static ApprovedPriceTable? FindTable(SheetGrid top)
    {
        var lastRow = Math.Min(top.LastRow, top.FirstRow + ScanRows - 1);
        for (var row = top.FirstRow; row <= lastRow; row++)
        {
            int? barcode = null;
            for (var column = top.FirstColumn; column <= top.LastColumn && barcode is null; column++)
            {
                if (BarcodeHeaders.Contains(TextUtils.NormalizeKey(top.Text(row, column))))
                {
                    barcode = column;
                }
            }

            if (barcode is null)
            {
                continue;
            }

            var price = FindPriceColumn(top, row);
            if (price is not null)
            {
                return new ApprovedPriceTable(
                    row, barcode.Value, price.Value, TextUtils.Normalize(top.Text(row, price.Value)));
            }
        }

        return null;
    }

    /// <summary>
    /// Колонка цены: сначала полное совпадение с «Розничная цена», потом с «Предложение УТЗ»,
    /// потом то же началом заголовка («Розничная цена, руб.»).
    /// </summary>
    private static int? FindPriceColumn(SheetGrid grid, int row)
    {
        foreach (var exact in new[] { true, false })
        {
            foreach (var header in PriceHeaders)
            {
                for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
                {
                    var text = TextUtils.NormalizeKey(grid.Text(row, column));
                    if (exact ? text == header : text.StartsWith(header, StringComparison.Ordinal))
                    {
                        return column;
                    }
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Дописывает цены таблицы в словарь по штрихкоду. Штрихкод, который уже есть,
    /// не перезаписывается: так файл поставки сильнее ближайших, а первая строка файла -
    /// следующих, как у ВПР.
    /// </summary>
    public static int Read(SheetGrid barcodes, int barcodeColumn, SheetGrid prices, int priceColumn, IDictionary<string, double> into)
    {
        var added = 0;
        for (var row = barcodes.FirstRow; row <= barcodes.LastRow; row++)
        {
            var barcode = InvoiceSheetReader.NormalizeBarcode(barcodes.Text(row, barcodeColumn));
            if (barcode.Length == 0 || into.ContainsKey(barcode) || !prices.Contains(row, priceColumn))
            {
                continue;
            }

            var cell = prices.Value(row, priceColumn);
            var price = CellError.IsError(cell) ? null : TextUtils.CellToDouble(cell);
            if (price is null)
            {
                continue;
            }

            into[barcode] = price.Value;
            added++;
        }

        return added;
    }
}
