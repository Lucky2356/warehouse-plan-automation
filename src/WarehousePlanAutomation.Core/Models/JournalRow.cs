namespace WarehousePlanAutomation.Core.Models;

/// <summary>
/// Строка листа «Журнал заказов на отгрузку».
/// <see cref="Order"/> хранит физический порядок строк: журнал не сортируется,
/// первое вхождение - это самая верхняя найденная строка.
/// </summary>
public sealed record JournalRow(int Order, int ExcelRow, string Comment, string Status, double? Percent);

/// <summary>Результат разбора журнала для одного номера загрузки.</summary>
public sealed record JournalOutcome(bool Found, bool SetInAssembly, double? PercentToSet)
{
    public static readonly JournalOutcome NotFound = new(false, false, null);
}
