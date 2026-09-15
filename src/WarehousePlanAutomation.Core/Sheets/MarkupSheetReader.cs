using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Чтение листа «Регламент наценок». Лист небольшой, читается целиком.
/// Строки без сектора или без обеих наценок пропускаются: внизу листа идут
/// «Порядок взаимодействия» и «Ответственные», к наценкам они отношения не имеют.
/// </summary>
public static class MarkupSheetReader
{
    public static HeaderMap ResolveHeaders(SheetGrid grid) =>
        HeaderResolver.Resolve(grid, PriceSchema.MarkupSheet, PriceSchema.Markup.Specs);

    public static IReadOnlyList<MarkupRule> Read(SheetGrid grid)
    {
        var headers = ResolveHeaders(grid);
        var sectorColumn = headers[PriceSchema.Markup.Sector];
        var sectorPlusColumn = headers[PriceSchema.Markup.SectorPlus];
        var plannedColumn = headers[PriceSchema.Markup.Planned];
        var minimumColumn = headers[PriceSchema.Markup.Minimum];

        var rules = new List<MarkupRule>();
        for (var row = headers.HeaderRow + 1; row <= grid.LastRow; row++)
        {
            var sector = TextUtils.Normalize(grid.Text(row, sectorColumn));
            if (sector.Length == 0)
            {
                continue;
            }

            var planned = grid.Number(row, plannedColumn);
            var minimum = grid.Number(row, minimumColumn);
            if (planned is null && minimum is null)
            {
                continue;
            }

            rules.Add(new MarkupRule(
                sector,
                TextUtils.Normalize(grid.Text(row, sectorPlusColumn)),
                planned,
                minimum));
        }

        return rules;
    }
}
