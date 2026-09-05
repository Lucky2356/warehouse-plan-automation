using WarehousePlanAutomation.Core.Models;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Номер, который нужно записать в колонку «№».</summary>
public sealed record NumberAssignment(int ExcelRow, int Number);

/// <summary>
/// Сквозная нумерация строк блоков реальных заказов.
///
/// Порядок строк программа не меняет: план ведёт аналитик, и её расстановка -
/// это решение, а не побочный результат правил. Нумерация только приводит колонку «№»
/// в соответствие с тем порядком, который сложился на листе после добавления
/// и удаления строк.
/// </summary>
public static class PlanNumberingBuilder
{
    public static IReadOnlyList<NumberAssignment> Build(PlanLayout plan)
    {
        var numbers = new List<NumberAssignment>();

        foreach (var section in plan.OrderSections)
        {
            var number = 1;
            foreach (var row in section.DataRows)
            {
                numbers.Add(new NumberAssignment(row.ExcelRow, number++));
            }
        }

        return numbers;
    }
}
