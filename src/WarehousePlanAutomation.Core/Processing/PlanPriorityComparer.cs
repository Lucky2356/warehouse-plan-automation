using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Строка блока «План» в виде, пригодном для приоритизации.</summary>
public sealed record PlanSortItem(
    int OriginalIndex,
    bool IsStorageAcceptanceRow,
    bool IsAutoHubRow,
    bool IsUrgent,
    double? NetworkDate,
    SetMonoKind SetMono);

/// <summary>
/// Приоритизация строк внутри блока.
///
/// Порядок сравнения:
/// 1) «Приемка на хранилище», затем «автозаказы для хабов» (только для блока «все группы»);
/// 2) срочные заказы выше несрочных;
/// 3) «Дата в сети» по возрастанию (пустая дата - в конец категории, значение не додумывается);
/// 4) при одинаковой дате «СЕТ» выше «МОНО»;
/// 5) при полном равенстве сохраняется исходный взаимный порядок.
/// </summary>
public sealed class PlanPriorityComparer : IComparer<PlanSortItem>
{
    private readonly bool _honorSpecialRows;

    public PlanPriorityComparer(bool honorSpecialRows)
    {
        _honorSpecialRows = honorSpecialRows;
    }

    public int Compare(PlanSortItem? x, PlanSortItem? y)
    {
        if (ReferenceEquals(x, y))
        {
            return 0;
        }

        if (x is null)
        {
            return -1;
        }

        if (y is null)
        {
            return 1;
        }

        if (_honorSpecialRows)
        {
            var specialCompare = SpecialRank(x).CompareTo(SpecialRank(y));
            if (specialCompare != 0)
            {
                return specialCompare;
            }
        }

        var urgencyCompare = UrgencyRank(x).CompareTo(UrgencyRank(y));
        if (urgencyCompare != 0)
        {
            return urgencyCompare;
        }

        var dateCompare = CompareDates(x.NetworkDate, y.NetworkDate);
        if (dateCompare != 0)
        {
            return dateCompare;
        }

        var setMonoCompare = ((int)x.SetMono).CompareTo((int)y.SetMono);
        if (setMonoCompare != 0)
        {
            return setMonoCompare;
        }

        return x.OriginalIndex.CompareTo(y.OriginalIndex);
    }

    private static int SpecialRank(PlanSortItem item)
    {
        if (item.IsStorageAcceptanceRow)
        {
            return 0;
        }

        return item.IsAutoHubRow ? 1 : 2;
    }

    private static int UrgencyRank(PlanSortItem item) => item.IsUrgent ? 0 : 1;

    private static int CompareDates(double? left, double? right)
    {
        if (left is null && right is null)
        {
            return 0;
        }

        if (left is null)
        {
            return 1;
        }

        if (right is null)
        {
            return -1;
        }

        return left.Value.CompareTo(right.Value);
    }
}
