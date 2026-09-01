using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>Чтение листа «Журнал заказов на отгрузку». Лист только читается и никогда не сортируется.</summary>
public sealed class JournalSheet
{
    public JournalSheet(HeaderMap headers, IReadOnlyList<JournalRow> rows)
    {
        Headers = headers;
        Rows = rows;
    }

    public HeaderMap Headers { get; }

    public IReadOnlyList<JournalRow> Rows { get; }
}

public static class JournalSheetReader
{
    public static HeaderMap ResolveHeaders(SheetGrid grid) =>
        HeaderResolver.Resolve(grid, SheetSchema.JournalSheet, SheetSchema.Journal.Specs);

    /// <summary>
    /// Разбирает строки журнала из прочитанного фрагмента листа.
    /// Физический порядок строк сохраняется: он определяет первое вхождение заказа.
    /// </summary>
    public static void ReadRows(SheetGrid grid, HeaderMap headers, ICollection<JournalRow> destination, ref int order)
    {
        var commentColumn = headers[SheetSchema.Journal.Comment];
        var statusColumn = headers[SheetSchema.Journal.Status];
        var percentColumn = headers[SheetSchema.Journal.Percent];

        var firstRow = Math.Max(grid.FirstRow, headers.HeaderRow + 1);
        for (var row = firstRow; row <= grid.LastRow; row++)
        {
            var comment = TextUtils.Normalize(grid.Text(row, commentColumn));
            var status = TextUtils.Normalize(grid.Text(row, statusColumn));
            var percent = grid.Number(row, percentColumn);

            if (comment.Length == 0 && status.Length == 0 && percent is null)
            {
                continue;
            }

            destination.Add(new JournalRow(order++, row, comment, status, percent));
        }
    }

    public static JournalSheet Read(SheetGrid grid)
    {
        var headers = ResolveHeaders(grid);
        var rows = new List<JournalRow>();
        var order = 0;
        ReadRows(grid, headers, rows, ref order);
        return new JournalSheet(headers, rows);
    }
}
