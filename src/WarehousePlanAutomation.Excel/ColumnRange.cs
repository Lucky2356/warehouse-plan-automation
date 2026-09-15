namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Диапазон колонок таблицы листа. Вставка, удаление и перенос строк выполняются только
/// внутри него: если двигать строку целиком, вместе с ней уезжает всё, что стоит справа
/// от таблицы. На листе «План» это боковая сводка «норма в день / просрочка / в план».
/// </summary>
public readonly record struct ColumnRange(int First, int Last);
