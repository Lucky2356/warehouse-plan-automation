namespace WarehousePlanAutomation.Core.Models;

/// <summary>
/// Строка листа «Журнал заказов на отгрузку».
///
/// <see cref="Order"/> хранит физический порядок строк: журнал не сортируется.
///
/// <see cref="DocumentNumber"/> - номер документа магазина вида «З000-260355».
/// У сводных строк там стоит число (например «-21»), а их «Кол-во ед» уже включает
/// строки магазинов, поэтому в подсчёте выполнения такие строки не участвуют.
/// </summary>
public sealed record JournalRow(
    int Order,
    int ExcelRow,
    string Comment,
    string Status,
    double? Percent,
    string DocumentNumber = "",
    double? Quantity = null,
    double? ActualQuantity = null);

/// <summary>Результат разбора журнала для одного номера загрузки.</summary>
public sealed record JournalOutcome(bool Found, bool SetInAssembly, double? PercentToSet)
{
    public static readonly JournalOutcome NotFound = new(false, false, null);
}
