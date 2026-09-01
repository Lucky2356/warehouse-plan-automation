using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.Core.Models;

/// <summary>Блоки листа «План».</summary>
public enum PlanSectionKind
{
    AllGroups,
    Returns,
    StorageAcceptance,
    Marketplaces,
    Wholesale,
    InternetShop,
}

/// <summary>Строка данных листа «План».</summary>
public sealed class PlanRow
{
    public PlanRow(int excelRow, PlanSectionKind section)
    {
        ExcelRow = excelRow;
        Section = section;
    }

    public int ExcelRow { get; }

    public PlanSectionKind Section { get; }

    public string Supplies { get; init; } = string.Empty;

    public string Processing { get; init; } = string.Empty;

    public string Group { get; init; } = string.Empty;

    public string Comments { get; init; } = string.Empty;

    public string Status { get; init; } = string.Empty;

    public double? DocumentDate { get; init; }

    public double? Quantity { get; init; }

    public double? CompletionPercent { get; init; }

    public double? NetworkDate { get; init; }

    public long? LoadNumber { get; init; }

    /// <summary>Строка «Приемка на хранилище» в блоке «все группы».</summary>
    public bool IsStorageAcceptanceRow { get; init; }

    /// <summary>Строка «автозаказы для хабов» в блоке «все группы».</summary>
    public bool IsAutoHubRow { get; init; }

    /// <summary>Колонки, содержащие формулу. Ключ - имя колонки из <see cref="SheetSchema.Plan"/>.</summary>
    public IReadOnlySet<string> FormulaColumns { get; init; } = new HashSet<string>();

    /// <summary>Обычная строка заказа: не «Приемка на хранилище» и не «автозаказы для хабов».</summary>
    public bool IsOrderRow => !IsStorageAcceptanceRow && !IsAutoHubRow;
}

/// <summary>Блок листа «План»: строка-заголовок и её строки данных.</summary>
public sealed class PlanSection
{
    public PlanSection(PlanSectionKind kind, int headerRow, string? aggregateFormula, List<PlanRow> dataRows)
    {
        Kind = kind;
        HeaderRow = headerRow;
        AggregateFormula = aggregateFormula;
        DataRows = dataRows;
    }

    public PlanSectionKind Kind { get; }

    public int HeaderRow { get; }

    /// <summary>Формула в колонке «Количество единиц» строки-заголовка, если она там есть.</summary>
    public string? AggregateFormula { get; }

    public List<PlanRow> DataRows { get; }

    public int FirstDataRow => DataRows.Count == 0 ? HeaderRow + 1 : DataRows[0].ExcelRow;

    public int LastDataRow => DataRows.Count == 0 ? HeaderRow : DataRows[^1].ExcelRow;
}

/// <summary>Разобранная структура листа «План».</summary>
public sealed class PlanLayout
{
    public PlanLayout(
        HeaderMap headers,
        IReadOnlyList<PlanSection> sections,
        PlanRow? marketplaceFromStorage,
        PlanRow? marketplaceFromReturns,
        PlanRow? marketplaceFromSupplies)
    {
        Headers = headers;
        Sections = sections;
        MarketplaceFromStorage = marketplaceFromStorage;
        MarketplaceFromReturns = marketplaceFromReturns;
        MarketplaceFromSupplies = marketplaceFromSupplies;
    }

    public HeaderMap Headers { get; }

    public IReadOnlyList<PlanSection> Sections { get; }

    public PlanRow? MarketplaceFromStorage { get; }

    public PlanRow? MarketplaceFromReturns { get; }

    public PlanRow? MarketplaceFromSupplies { get; }

    public PlanSection? Section(PlanSectionKind kind) => Sections.FirstOrDefault(s => s.Kind == kind);

    /// <summary>Блоки реальных заказов: «все группы», «возвраты», «приемка на хранилище».</summary>
    public IEnumerable<PlanSection> OrderSections =>
        Sections.Where(s => s.Kind is PlanSectionKind.AllGroups
            or PlanSectionKind.Returns
            or PlanSectionKind.StorageAcceptance);

    public IEnumerable<PlanRow> AllDataRows => Sections.SelectMany(s => s.DataRows);

    public PlanRow? AutoHubRow => AllDataRows.FirstOrDefault(r => r.IsAutoHubRow);

    public PlanRow? StorageAcceptanceRow => AllDataRows.FirstOrDefault(r => r.IsStorageAcceptanceRow);
}
