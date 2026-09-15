using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Чтение справочников «Сводный прайс» и «Прайс по подразделениям».
///
/// Оба листа - сотни тысяч строк, а нужны из них два-три десятка: те, чьи штрихкоды
/// есть в инвойсе. Поэтому читатель принимает набор искомых ключей и запоминает только
/// совпавшие строки - книга целиком в памяти не оседает.
/// </summary>
public static class PriceListReader
{
    public static HeaderMap ResolveSummaryHeaders(SheetGrid grid) =>
        HeaderResolver.Resolve(grid, PriceSchema.SummaryPriceSheet, PriceSchema.SummaryPrice.Specs);

    public static HeaderMap ResolveDivisionHeaders(SheetGrid grid) =>
        HeaderResolver.Resolve(grid, PriceSchema.DivisionPriceSheet, PriceSchema.DivisionPrice.Specs);

    /// <summary>
    /// Собирает строки «Сводного прайса» по нужным штрихкодам.
    /// Побеждает первое вхождение: прайс отсортирован, и повторы - это история цен.
    /// </summary>
    public static void ReadSummary(
        SheetGrid grid,
        HeaderMap headers,
        IReadOnlySet<string> wantedBarcodes,
        Dictionary<string, PriceListRow> result)
    {
        var barcodeColumn = headers[PriceSchema.SummaryPrice.Barcode];
        var firstRow = Math.Max(headers.HeaderRow + 1, grid.FirstRow);

        for (var row = firstRow; row <= grid.LastRow; row++)
        {
            var barcode = InvoiceSheetReader.NormalizeBarcode(grid.Text(row, barcodeColumn));
            if (barcode.Length == 0 || !wantedBarcodes.Contains(barcode) || result.ContainsKey(barcode))
            {
                continue;
            }

            string Cell(string column) => TextUtils.Normalize(grid.Text(row, headers[column]));

            result[barcode] = new PriceListRow(
                barcode,
                Cell(PriceSchema.SummaryPrice.Code),
                Cell(PriceSchema.SummaryPrice.SupplierArticle),
                Cell(PriceSchema.SummaryPrice.Sector),
                Cell(PriceSchema.SummaryPrice.Group),
                Cell(PriceSchema.SummaryPrice.Name),
                Cell(PriceSchema.SummaryPrice.Article),
                Cell(PriceSchema.SummaryPrice.Color),
                Cell(PriceSchema.SummaryPrice.Season),
                Cell(PriceSchema.SummaryPrice.Size));
        }
    }

    /// <summary>Собирает подгруппы по кодам товара.</summary>
    public static void ReadSubgroups(
        SheetGrid grid,
        HeaderMap headers,
        IReadOnlySet<string> wantedCodes,
        Dictionary<string, string> result)
    {
        var codeColumn = headers[PriceSchema.DivisionPrice.Code];
        var subgroupColumn = headers[PriceSchema.DivisionPrice.Subgroup];
        var firstRow = Math.Max(headers.HeaderRow + 1, grid.FirstRow);

        for (var row = firstRow; row <= grid.LastRow; row++)
        {
            var code = TextUtils.Normalize(grid.Text(row, codeColumn));
            if (code.Length == 0 || !wantedCodes.Contains(code) || result.ContainsKey(code))
            {
                continue;
            }

            var subgroup = TextUtils.Normalize(grid.Text(row, subgroupColumn));
            if (subgroup.Length > 0)
            {
                result[code] = subgroup;
            }
        }
    }
}
