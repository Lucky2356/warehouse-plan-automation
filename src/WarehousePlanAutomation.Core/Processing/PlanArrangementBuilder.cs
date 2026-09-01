using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Перемещение строки Excel: строка <see cref="FromRow"/> вставляется перед строкой <see cref="ToRow"/>.</summary>
public sealed record RowMove(int FromRow, int ToRow);

/// <summary>Номер, который нужно записать в колонку «№».</summary>
public sealed record NumberAssignment(int ExcelRow, int Number);

/// <summary>Результат приоритизации и нумерации листа «План».</summary>
public sealed class PlanArrangement
{
    public PlanArrangement(IReadOnlyList<RowMove> moves, IReadOnlyList<NumberAssignment> numbers)
    {
        Moves = moves;
        Numbers = numbers;
    }

    public IReadOnlyList<RowMove> Moves { get; }

    public IReadOnlyList<NumberAssignment> Numbers { get; }
}

/// <summary>
/// Приоритизация строк блоков «все группы» и «приемка на хранилище» и сквозная нумерация
/// блоков реальных заказов. Блок «возвраты» сохраняет исходный порядок.
/// </summary>
public static class PlanArrangementBuilder
{
    public static PlanArrangement Build(PlanLayout plan)
    {
        var moves = new List<RowMove>();

        foreach (var kind in new[] { PlanSectionKind.AllGroups, PlanSectionKind.StorageAcceptance })
        {
            var section = plan.Section(kind);
            if (section is null || section.DataRows.Count < 2)
            {
                continue;
            }

            moves.AddRange(BuildSectionMoves(section, honorSpecialRows: kind == PlanSectionKind.AllGroups));
        }

        var numbers = new List<NumberAssignment>();
        foreach (var section in plan.OrderSections)
        {
            var number = 1;
            foreach (var row in section.DataRows)
            {
                numbers.Add(new NumberAssignment(row.ExcelRow, number++));
            }
        }

        return new PlanArrangement(moves, numbers);
    }

    /// <summary>
    /// Строит последовательность перемещений, приводящую физический порядок строк блока
    /// к целевому. Каждое перемещение соответствует операции Excel «вырезать строку и
    /// вставить её перед другой строкой», поэтому позиции пересчитываются после каждого шага.
    /// </summary>
    public static IReadOnlyList<RowMove> BuildSectionMoves(PlanSection section, bool honorSpecialRows)
    {
        var items = new List<PlanSortItem>(section.DataRows.Count);
        for (var i = 0; i < section.DataRows.Count; i++)
        {
            var row = section.DataRows[i];
            items.Add(new PlanSortItem(
                i,
                row.IsStorageAcceptanceRow,
                row.IsAutoHubRow,
                OrderTextRules.IsUrgent(row.Supplies),
                row.NetworkDate,
                OrderTextRules.DetectSetMono(row.Supplies)));
        }

        var desired = items.ToList();
        desired.Sort(new PlanPriorityComparer(honorSpecialRows));

        var firstRow = section.FirstDataRow;
        var current = items.Select(i => i.OriginalIndex).ToList();
        var moves = new List<RowMove>();

        for (var target = 0; target < desired.Count; target++)
        {
            var wanted = desired[target].OriginalIndex;
            if (current[target] == wanted)
            {
                continue;
            }

            var source = current.IndexOf(wanted, target + 1);
            if (source < 0)
            {
                continue;
            }

            moves.Add(new RowMove(firstRow + source, firstRow + target));
            current.Insert(target, wanted);
            current.RemoveAt(source + 1);
        }

        return moves;
    }
}
