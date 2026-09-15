namespace WarehousePlanAutomation.Core.Models;

/// <summary>Строка товара в инвойсе.</summary>
public sealed record InvoiceLine(
    int ExcelRow,
    string Barcode,
    double Quantity,
    double Price,
    double Amount);

/// <summary>Разобранный лист инвойса.</summary>
public sealed class InvoiceSheet
{
    public InvoiceSheet(string sheetName, string supplyNumber, IReadOnlyList<InvoiceLine> lines)
    {
        SheetName = sheetName;
        SupplyNumber = supplyNumber;
        Lines = lines;
    }

    public string SheetName { get; }

    /// <summary>Номер поставки из шапки, например «431-329». Пусто, если найти не удалось.</summary>
    public string SupplyNumber { get; }

    public IReadOnlyList<InvoiceLine> Lines { get; }

    public double TotalQuantity => Lines.Sum(line => line.Quantity);

    public double TotalAmount => Lines.Sum(line => line.Amount);
}

/// <summary>Строка справочника «Сводный прайс», найденная по штрихкоду.</summary>
public sealed record PriceListRow(
    string Barcode,
    string Code,
    string Model,
    string Sector,
    string Group,
    string Name,
    string Article,
    string Color,
    string Season,
    string Size);

/// <summary>Строка «Регламента наценок».</summary>
public sealed record MarkupRule(string Sector, string SectorPlus, double? Planned, double? Minimum);

/// <summary>
/// Строка листа «link»: один магазин для одного АЦР.
/// </summary>
/// <param name="StoreGrade">«Группа_ам» - грейд магазина: A, B, C1, …, O120, O140.</param>
/// <param name="Concept">«Концепт»: МАГАЗИН, МАГАЗИН-ХАБ, ОСТРОВ. Пусто, если колонки нет.</param>
public sealed record LinkRow(
    string Acr,
    double? DenyGoodsDivision,
    double? DenyGoodsAll,
    double? Included,
    string StoreGrade,
    double? DateStart,
    double? DateEnd,
    string Concept = "")
{
    /// <summary>
    /// Магазину этот АЦР положен: «Вкл» = 1 и обе «Deny goods» = 0. Пустая ячейка
    /// ничего не запрещает - запрет всегда записан явно, нулём во «Вкл» или единицей в «Deny goods».
    /// </summary>
    public bool IsAllowed =>
        Included is null or 1d && DenyGoodsDivision is null or 0d && DenyGoodsAll is null or 0d;
}

/// <summary>Магазины одного грейда в «link» для одного АЦР: сколько их и скольким АЦР положен.</summary>
public sealed record LinkGrade(string Grade, int Stores, int AllowedStores)
{
    public bool IsFullyDenied => AllowedStores == 0;

    public bool IsPartlyDenied => AllowedStores > 0 && AllowedStores < Stores;
}

/// <summary>
/// Сведения по одному АЦР, собранные из всех его строк «link».
/// </summary>
public sealed record LinkEntry(
    string Acr,
    IReadOnlyList<LinkGrade> Grades,
    double? DateStart,
    double? DateEnd,
    int Rows,
    int AllowedRows);
