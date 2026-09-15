using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Разбор листа «План»: строка заголовков, блоки, строки данных, особые и агрегатные строки.
/// Ни одна строка и ни одна колонка не задаются фиксированными координатами.
/// </summary>
public static class PlanSheetReader
{
    private static readonly Dictionary<string, PlanSectionKind> SectionKeys = new(StringComparer.Ordinal)
    {
        [SheetSchema.Plan.SectionAllGroups] = PlanSectionKind.AllGroups,
        [SheetSchema.Plan.SectionReturns] = PlanSectionKind.Returns,
        [SheetSchema.Plan.SectionStorageAcceptance] = PlanSectionKind.StorageAcceptance,
        [SheetSchema.Plan.SectionMarketplaces] = PlanSectionKind.Marketplaces,
        [SheetSchema.Plan.SectionWholesale] = PlanSectionKind.Wholesale,
        [SheetSchema.Plan.SectionInternetShop] = PlanSectionKind.InternetShop,
    };

    public static PlanLayout Read(SheetGrid grid)
    {
        var headers = HeaderResolver.Resolve(grid, SheetSchema.PlanSheet, SheetSchema.Plan.Specs);
        var columns = headers.Columns;
        var firstColumn = columns.Values.Min();
        var lastColumn = columns.Values.Max();

        var numberColumn = headers[SheetSchema.Plan.Number];
        var quantityColumn = headers[SheetSchema.Plan.Quantity];
        var commentsColumn = headers[SheetSchema.Plan.Comments];
        var suppliesColumn = headers[SheetSchema.Plan.Supplies];

        var sections = new List<PlanSection>();
        PlanSection? current = null;

        for (var row = headers.HeaderRow + 1; row <= grid.LastRow; row++)
        {
            var sectionKey = TextUtils.NormalizeKey(grid.Text(row, numberColumn));
            if (SectionKeys.TryGetValue(sectionKey, out var kind))
            {
                var formula = grid.Formula(row, quantityColumn);
                current = new PlanSection(
                    kind,
                    row,
                    formula is not null && formula.StartsWith("=", StringComparison.Ordinal) ? formula : null,
                    new List<PlanRow>());
                sections.Add(current);
                continue;
            }

            if (current is null || IsBlank(grid, row, firstColumn, lastColumn))
            {
                continue;
            }

            current.DataRows.Add(ReadRow(grid, headers, row, current.Kind, suppliesColumn));
        }

        var marketplaces = sections.FirstOrDefault(s => s.Kind == PlanSectionKind.Marketplaces);
        PlanRow? FindMarketplaceRow(string marker) => marketplaces?.DataRows
            .FirstOrDefault(r => TextUtils.EqualsKey(r.Comments, marker));

        var layout = new PlanLayout(
            headers,
            sections,
            FindMarketplaceRow(SheetSchema.Plan.MarketplaceFromStorage),
            FindMarketplaceRow(SheetSchema.Plan.MarketplaceFromReturns),
            FindMarketplaceRow(SheetSchema.Plan.MarketplaceFromSupplies));

        Validate(layout);
        return layout;
    }

    private static PlanRow ReadRow(
        SheetGrid grid,
        HeaderMap headers,
        int row,
        PlanSectionKind kind,
        int suppliesColumn)
    {
        var formulaColumns = new HashSet<string>(StringComparer.Ordinal);
        foreach (var pair in headers.Columns)
        {
            if (grid.HasFormula(row, pair.Value))
            {
                formulaColumns.Add(pair.Key);
            }
        }

        var supplies = TextUtils.Normalize(grid.Text(row, suppliesColumn));
        var isAllGroups = kind == PlanSectionKind.AllGroups;

        return new PlanRow(row, kind)
        {
            Supplies = supplies,
            Processing = TextUtils.Normalize(grid.Text(row, headers[SheetSchema.Plan.Processing])),
            Group = TextUtils.Normalize(grid.Text(row, headers[SheetSchema.Plan.Group])),
            Comments = TextUtils.Normalize(grid.Text(row, headers[SheetSchema.Plan.Comments])),
            Status = TextUtils.Normalize(grid.Text(row, headers[SheetSchema.Plan.Status])),
            DocumentDate = grid.Number(row, headers[SheetSchema.Plan.DocumentDate]),
            Quantity = grid.Number(row, headers[SheetSchema.Plan.Quantity]),
            CompletionPercent = grid.Number(row, headers[SheetSchema.Plan.CompletionPercent]),
            NetworkDate = grid.Number(row, headers[SheetSchema.Plan.NetworkDate]),
            LoadNumber = ReadLoadNumber(grid, row, headers[SheetSchema.Plan.LoadNumber]),
            IsStorageAcceptanceRow = isAllGroups &&
                                     TextUtils.StartsWithKey(supplies, SheetSchema.Plan.StorageAcceptanceRow),
            IsAutoHubRow = isAllGroups && TextUtils.StartsWithKey(supplies, SheetSchema.Plan.AutoHubRow),
            FormulaColumns = formulaColumns,
        };
    }

    private static long? ReadLoadNumber(SheetGrid grid, int row, int column)
    {
        var number = grid.Number(row, column);
        if (number is not null && number.Value >= 0 && Math.Abs(number.Value - Math.Round(number.Value)) < 1e-9)
        {
            var rounded = (long)Math.Round(number.Value);
            return rounded == 0 ? null : rounded;
        }

        var text = TextUtils.Normalize(grid.Text(row, column));
        var digits = new string(text.Where(char.IsDigit).ToArray());
        return digits.Length > 0 && long.TryParse(digits, out var parsed) ? parsed : null;
    }

    private static bool IsBlank(SheetGrid grid, int row, int firstColumn, int lastColumn)
    {
        for (var column = firstColumn; column <= lastColumn; column++)
        {
            if (TextUtils.Normalize(grid.Text(row, column)).Length > 0)
            {
                return false;
            }
        }

        return true;
    }

    private static void Validate(PlanLayout layout)
    {
        var problems = new List<string>();

        void RequireSection(PlanSectionKind kind, string name)
        {
            if (layout.Section(kind) is null)
            {
                problems.Add("на листе «" + SheetSchema.PlanSheet + "» не найдена строка «" + name + "»");
            }
        }

        RequireSection(PlanSectionKind.AllGroups, SheetSchema.Plan.SectionAllGroups);
        RequireSection(PlanSectionKind.Returns, SheetSchema.Plan.SectionReturns);
        RequireSection(PlanSectionKind.StorageAcceptance, SheetSchema.Plan.SectionStorageAcceptance);
        RequireSection(PlanSectionKind.Marketplaces, SheetSchema.Plan.SectionMarketplaces);
        RequireSection(PlanSectionKind.Wholesale, SheetSchema.Plan.SectionWholesale);
        RequireSection(PlanSectionKind.InternetShop, SheetSchema.Plan.SectionInternetShop);

        if (layout.MarketplaceFromStorage is null)
        {
            problems.Add("в блоке «заказы МП» не найдена строка «" + SheetSchema.Plan.MarketplaceFromStorage + "»");
        }

        if (layout.MarketplaceFromReturns is null)
        {
            problems.Add("в блоке «заказы МП» не найдена строка «" + SheetSchema.Plan.MarketplaceFromReturns + "»");
        }

        if (layout.MarketplaceFromSupplies is null)
        {
            problems.Add("в блоке «заказы МП» не найдена строка «" + SheetSchema.Plan.MarketplaceFromSupplies + "»");
        }

        if (layout.StorageAcceptanceRow is null)
        {
            problems.Add("в блоке «все группы» не найдена строка «Приемка на хранилище»");
        }

        if (layout.AutoHubRow is null)
        {
            problems.Add("в блоке «все группы» не найдена строка «автозаказы для хабов»");
        }

        if (problems.Count > 0)
        {
            throw new WorkbookValidationException(problems);
        }
    }
}
