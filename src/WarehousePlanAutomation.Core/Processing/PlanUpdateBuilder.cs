using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Расчёт изменений листа «План» по вчерашнему плану, сегодняшней выгрузке и журналу.
/// Класс не знает ничего про Excel и полностью покрывается тестами.
/// </summary>
public static class PlanUpdateBuilder
{
    public static PlanStructuralUpdate Build(
        PlanLayout plan,
        OrderClassificationResult orders,
        IReadOnlyList<JournalRow> journal,
        DateTime today)
    {
        // Снимок вчерашнего плана снимается ДО добавления новых заказов.
        var yesterdayLoadNumbers = new HashSet<long>(
            plan.AllDataRows.Where(r => r.LoadNumber.HasValue).Select(r => r.LoadNumber!.Value));

        var todayGroups = orders.Groups.ToDictionary(g => g.LoadNumber);

        var newRows = BuildNewRows(plan, orders.Groups, yesterdayLoadNumbers);
        var orderUpdates = BuildOrderUpdates(plan, todayGroups, newRows, journal);
        var rowsToDelete = BuildPlaceholderDeletions(plan, orders.Groups);
        var aggregates = BuildAggregates(orders, plan, today);

        return new PlanStructuralUpdate(rowsToDelete, newRows, orderUpdates, aggregates);
    }

    private static IReadOnlyList<NewPlanRowSpec> BuildNewRows(
        PlanLayout plan,
        IReadOnlyList<TodayOrderGroup> groups,
        HashSet<long> yesterdayLoadNumbers)
    {
        _ = plan;
        var result = new List<NewPlanRowSpec>();
        foreach (var group in groups)
        {
            if (yesterdayLoadNumbers.Contains(group.LoadNumber))
            {
                continue;
            }

            var supplies = LoadNumberParser.ExtractSuppliesText(group.Comment);
            var section = OrderTextRules.IsReturn(group.Comment)
                ? PlanSectionKind.Returns
                : PlanSectionKind.AllGroups;

            // Новая поставка: номера загрузки вчера не было, а в тексте есть номер поставки.
            var processing = ShipmentCodeParser.ContainsCode(group.Comment)
                ? OrderTextRules.CrossDockProcessing
                : string.Empty;

            result.Add(new NewPlanRowSpec(
                section,
                supplies,
                processing,
                OrderTextRules.LoadedComment,
                NormalizeDocumentDate(group.DocumentDate),
                group.LoadNumber));
        }

        return result;
    }

    /// <summary>
    /// «Дата документа» в выгрузке хранится вместе со временем, а в «Плане» - только дата.
    /// Отбрасывание времени сохраняет целые значения в формуле «Дней в работе».
    /// </summary>
    private static double? NormalizeDocumentDate(double? value) =>
        value is null ? null : Math.Floor(value.Value);

    private static IReadOnlyList<OrderRowUpdate> BuildOrderUpdates(
        PlanLayout plan,
        IReadOnlyDictionary<long, TodayOrderGroup> todayGroups,
        IReadOnlyList<NewPlanRowSpec> newRows,
        IReadOnlyList<JournalRow> journal)
    {
        var loadNumbers = new List<long>();
        var seen = new HashSet<long>();

        foreach (var row in plan.OrderSections.SelectMany(s => s.DataRows))
        {
            if (row.IsOrderRow && row.LoadNumber.HasValue && seen.Add(row.LoadNumber.Value))
            {
                loadNumbers.Add(row.LoadNumber.Value);
            }
        }

        foreach (var spec in newRows)
        {
            if (seen.Add(spec.LoadNumber))
            {
                loadNumbers.Add(spec.LoadNumber);
            }
        }

        var updates = new List<OrderRowUpdate>(loadNumbers.Count);
        foreach (var loadNumber in loadNumbers)
        {
            // Заказ, которого сегодня нет среди обычных заказов, обнуляется по количеству.
            var quantity = todayGroups.TryGetValue(loadNumber, out var group) ? group.Quantity : 0d;

            var outcome = JournalEvaluator.Evaluate(loadNumber, journal);
            if (!outcome.Found)
            {
                updates.Add(new OrderRowUpdate(loadNumber, 0d, null, null));
                continue;
            }

            var status = outcome.SetInAssembly ? OrderTextRules.InAssemblyStatus : null;
            updates.Add(new OrderRowUpdate(loadNumber, quantity, status, outcome.PercentToSet));
        }

        return updates;
    }

    private static IReadOnlyList<int> BuildPlaceholderDeletions(
        PlanLayout plan,
        IReadOnlyList<TodayOrderGroup> groups)
    {
        var realCodes = new HashSet<string>(StringComparer.Ordinal);
        foreach (var group in groups)
        {
            foreach (var code in ShipmentCodeParser.ExtractCodes(group.Comment))
            {
                realCodes.Add(code);
            }
        }

        if (realCodes.Count == 0)
        {
            return Array.Empty<int>();
        }

        var rows = new List<int>();
        foreach (var row in plan.OrderSections.SelectMany(s => s.DataRows))
        {
            if (!row.IsOrderRow || !OrderTextRules.IsPlaceholderComment(row.Comments))
            {
                continue;
            }

            if (ShipmentCodeParser.ExtractCodes(row.Supplies).Any(realCodes.Contains))
            {
                rows.Add(row.ExcelRow);
            }
        }

        return rows;
    }

    private static IReadOnlyList<AggregateUpdate> BuildAggregates(
        OrderClassificationResult orders,
        PlanLayout plan,
        DateTime today)
    {
        var result = new List<AggregateUpdate>
        {
            new(PlanAggregateTarget.Wholesale, orders.Total(OrderCategory.Wholesale)),
            new(PlanAggregateTarget.InternetShop, orders.Total(OrderCategory.InternetShop)),
            new(PlanAggregateTarget.MarketplaceFromReturns, orders.Total(OrderCategory.MarketplaceReturns)),
            new(PlanAggregateTarget.MarketplaceFromSupplies, orders.Total(OrderCategory.MarketplaceSupplies)),
            new(PlanAggregateTarget.MarketplaceFromStorage, orders.Total(OrderCategory.MarketplaceStorage)),
        };

        var autoHubQuantity = AutoHubQuantityCalculator.GetQuantity(today.DayOfWeek);
        if (autoHubQuantity.HasValue && plan.AutoHubRow is not null)
        {
            result.Add(new AggregateUpdate(PlanAggregateTarget.AutoHub, autoHubQuantity.Value));
        }

        return result;
    }
}
