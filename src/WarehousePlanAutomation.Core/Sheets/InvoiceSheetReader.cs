using System.Text.RegularExpressions;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Разбор листа инвойса. Строка заголовков ищется по «штрих-код / количество / цена /
/// сумма»; сразу под ней в реальных инвойсах стоит вторая строка заголовков на английском,
/// поэтому строкой товара считается только та, где есть штрихкод и «количество» - число.
/// Итоговая строка «итого» отсекается тем же правилом: штрихкода в ней нет.
/// </summary>
public static class InvoiceSheetReader
{
    /// <summary>
    /// Заголовки товарной таблицы стоят глубоко: над ними шапка инвойса с продавцом,
    /// покупателем, условиями поставки и портами. В разобранном инвойсе это строка 23,
    /// поэтому просматривается заметно больше строк, чем на обычном листе.
    /// </summary>
    public const int HeaderScanRows = 40;

    public static HeaderMap ResolveHeaders(SheetGrid grid) =>
        HeaderResolver.Resolve(grid, PriceSchema.InvoiceSheet, PriceSchema.Invoice.Specs, HeaderScanRows);

    /// <param name="fileName">
    /// Название книги: если в шапке номера поставки нет, он ищется по листу и файлу
    /// (см. <see cref="SupplyFromNames"/>).
    /// </param>
    public static InvoiceSheet Read(SheetGrid grid, string sheetName, string? fileName = null)
    {
        var headers = ResolveHeaders(grid);
        var lines = new List<InvoiceLine>();
        ReadLines(grid, headers, lines);

        var supply = FindSupplyNumber(grid, headers);
        if (supply.Length == 0)
        {
            supply = SupplyFromNames(sheetName, fileName);
        }

        return new InvoiceSheet(sheetName, supply, lines);
    }

    /// <summary>
    /// Номер поставки по названиям: лист «Invoice-341» в файле «431-338, 431-341 Бижутерия»
    /// - это поставка «431-341». Из названия листа берётся последнее число, и из номеров
    /// в названии файла выбирается тот, который им кончается. Не нашлось или подошло
    /// несколько - пусто: гадать нельзя.
    /// </summary>
    public static string SupplyFromNames(string? sheetName, string? fileName)
    {
        var numbers = Regex.Matches(TextUtils.Normalize(sheetName), @"\d+");
        if (numbers.Count == 0)
        {
            return string.Empty;
        }

        var tail = numbers[^1].Value.TrimStart('0');
        var matches = ShipmentCodeParser.ExtractCodes(fileName)
            .Where(code => string.Equals(code[(code.IndexOf('-') + 1)..].TrimStart('0'), tail, StringComparison.Ordinal))
            .ToList();

        return matches.Count == 1 ? matches[0] : string.Empty;
    }

    public static void ReadLines(SheetGrid grid, HeaderMap headers, List<InvoiceLine> lines)
    {
        var barcodeColumn = headers[PriceSchema.Invoice.Barcode];
        var quantityColumn = headers[PriceSchema.Invoice.Quantity];
        var priceColumn = headers[PriceSchema.Invoice.Price];
        var amountColumn = headers[PriceSchema.Invoice.Amount];

        var firstRow = Math.Max(headers.HeaderRow + 1, grid.FirstRow);
        for (var row = firstRow; row <= grid.LastRow; row++)
        {
            var barcode = NormalizeBarcode(grid.Text(row, barcodeColumn));
            if (barcode.Length == 0)
            {
                continue;
            }

            var quantity = grid.Number(row, quantityColumn);
            if (quantity is null)
            {
                continue;
            }

            lines.Add(new InvoiceLine(
                row,
                barcode,
                quantity.Value,
                grid.Number(row, priceColumn) ?? 0d,
                grid.Number(row, amountColumn) ?? 0d));
        }
    }

    /// <summary>
    /// Штрихкод сравнивается как текст: длинное число, прочитанное как double,
    /// иначе превратилось бы в «1,9E+09» и перестало совпадать с прайсом.
    /// </summary>
    public static string NormalizeBarcode(string? value) => TextUtils.Normalize(value).Replace(" ", string.Empty);

    /// <summary>
    /// Номер поставки берётся из шапки: под подписью «Инвойс № и дата» стоит значение
    /// вида «431-329». Разбирается тем же правилом, что и номера поставок в плане склада.
    /// </summary>
    private static string FindSupplyNumber(SheetGrid grid, HeaderMap headers)
    {
        for (var row = grid.FirstRow; row < headers.HeaderRow; row++)
        {
            for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
            {
                if (!TextUtils.ContainsKey(grid.Text(row, column), PriceSchema.Invoice.SupplyLabel))
                {
                    continue;
                }

                var codes = ShipmentCodeParser.ExtractCodes(grid.Text(row + 1, column));
                if (codes.Count > 0)
                {
                    return codes[0];
                }
            }
        }

        return string.Empty;
    }
}
