using System.Globalization;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Формулы книги приемки, которые программа вписывает сама.
///
/// «В приемку» на листах «А2, А3» и «МП» - сколько по коду строки решено допоставить:
/// =СУММЕСЛИ(итог!G:G;D2;итог!S:S). На неё опирается «Учитывать»: количество набирается
/// с меньших тар, пока не покроет это число. Формула подтверждена аналитиком; буквы
/// колонок берутся из заголовков, а не зашиваются.
/// </summary>
public static class ReceivingFormulas
{
    /// <summary>Формула для записи через Range.Formula: английские имена, запятые.</summary>
    public static string ToReceive(
        string summarySheet, int summaryCodeColumn, int summaryRestockColumn, int storageCodeColumn, int row) =>
        "=SUMIF(" + Arguments(summarySheet, summaryCodeColumn, summaryRestockColumn, storageCodeColumn, row, ",") + ")";

    /// <summary>Та же формула так, как её видит аналитик в русском Excel - для замечаний.</summary>
    public static string ToReceiveLocal(
        string summarySheet, int summaryCodeColumn, int summaryRestockColumn, int storageCodeColumn, int row) =>
        "=СУММЕСЛИ(" + Arguments(summarySheet, summaryCodeColumn, summaryRestockColumn, storageCodeColumn, row, ";") + ")";

    private static string Arguments(
        string summarySheet, int summaryCodeColumn, int summaryRestockColumn, int storageCodeColumn, int row, string separator)
    {
        var sheet = SheetPrefix(summarySheet);
        var code = ExcelColumn.ToLetters(summaryCodeColumn);
        var restock = ExcelColumn.ToLetters(summaryRestockColumn);
        return sheet + code + ":" + code + separator +
               ExcelColumn.ToLetters(storageCodeColumn) + row.ToString(CultureInfo.InvariantCulture) + separator +
               sheet + restock + ":" + restock;
    }

    /// <summary>Имя листа в ссылке: в кавычках только тогда, когда без них нельзя.</summary>
    private static string SheetPrefix(string sheet) =>
        sheet.All(ch => char.IsLetterOrDigit(ch) || ch == '_')
            ? sheet + "!"
            : "'" + sheet.Replace("'", "''", StringComparison.Ordinal) + "'!";
}
