using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>Чтение листа «Заказы на отгрузку» в память.</summary>
public sealed class OrdersSheet
{
    public OrdersSheet(HeaderMap headers, IReadOnlyList<OrderRow> rows)
    {
        Headers = headers;
        Rows = rows;
    }

    public HeaderMap Headers { get; }

    public IReadOnlyList<OrderRow> Rows { get; }
}

public static class OrdersSheetReader
{
    public static HeaderMap ResolveHeaders(SheetGrid grid) =>
        HeaderResolver.Resolve(grid, SheetSchema.OrdersSheet, SheetSchema.Orders.Specs);

    /// <summary>
    /// Разбирает строки данных из прочитанного фрагмента листа.
    /// Позволяет читать большие выгрузки по частям, не держа весь лист в памяти.
    /// </summary>
    public static void ReadRows(SheetGrid grid, HeaderMap headers, ICollection<OrderRow> destination)
    {
        var divisionColumn = headers[SheetSchema.Orders.Division];
        var commentColumn = headers[SheetSchema.Orders.Comment];
        var differenceColumn = headers[SheetSchema.Orders.DifferenceUnits];
        var documentDateColumn = headers[SheetSchema.Orders.DocumentDate];

        var firstRow = Math.Max(grid.FirstRow, headers.HeaderRow + 1);
        for (var row = firstRow; row <= grid.LastRow; row++)
        {
            var division = TextUtils.Normalize(grid.Text(row, divisionColumn));
            var comment = TextUtils.Normalize(grid.Text(row, commentColumn));
            var difference = grid.Number(row, differenceColumn);
            var documentDate = grid.Number(row, documentDateColumn);

            if (division.Length == 0 && comment.Length == 0 && difference is null && documentDate is null)
            {
                continue;
            }

            destination.Add(new OrderRow(row, division, comment, difference ?? 0d, documentDate));
        }
    }

    public static OrdersSheet Read(SheetGrid grid)
    {
        var headers = ResolveHeaders(grid);
        var rows = new List<OrderRow>();
        ReadRows(grid, headers, rows);
        return new OrdersSheet(headers, rows);
    }
}
