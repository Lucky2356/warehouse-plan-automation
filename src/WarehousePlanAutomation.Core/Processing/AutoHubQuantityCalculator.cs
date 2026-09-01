namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// «Автозаказы для хабов»: значение «Количество единиц» по текущему дню недели.
/// Собственный интервал до пятницы не вычисляется - используется фиксированная таблица из ТЗ.
/// </summary>
public static class AutoHubQuantityCalculator
{
    /// <summary>Возвращает значение или null, если в этот день значение менять не нужно.</summary>
    public static double? GetQuantity(DayOfWeek dayOfWeek) => dayOfWeek switch
    {
        DayOfWeek.Monday => 24000d,
        DayOfWeek.Tuesday => 18000d,
        DayOfWeek.Wednesday => 12000d,
        DayOfWeek.Thursday => 6000d,
        DayOfWeek.Friday => 42000d,
        _ => null,
    };
}
