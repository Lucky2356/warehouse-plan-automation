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

        var (newRows, plannedMatches) = BuildNewRows(plan, orders.Groups, yesterdayLoadNumbers);
        var orderUpdates = BuildOrderUpdates(plan, todayGroups, newRows, plannedMatches, journal);
        var rowsToDelete = BuildPlaceholderDeletions(plan, orders.Groups, plannedMatches);
        var aggregates = BuildAggregates(orders, plan, today);

        return new PlanStructuralUpdate(rowsToDelete, newRows, orderUpdates, aggregates, plannedMatches);
    }

    private static (IReadOnlyList<NewPlanRowSpec> NewRows, IReadOnlyList<PlannedRowMatch> Matches) BuildNewRows(
        PlanLayout plan,
        IReadOnlyList<TodayOrderGroup> groups,
        HashSet<long> yesterdayLoadNumbers)
    {
        // Строки «Плана», заведённые заранее: поставка запланирована, но номера загрузки
        // у неё ещё нет. Ключ - текст «Поставки», по нему заказ и узнаётся, когда приходит.
        var plannedRows = new Dictionary<string, PlanRow>(StringComparer.Ordinal);
        foreach (var row in plan.OrderSections.SelectMany(s => s.DataRows))
        {
            if (!row.IsOrderRow || row.LoadNumber.HasValue || row.Supplies.Length == 0)
            {
                continue;
            }

            var key = TextUtils.NormalizeKey(row.Supplies);
            if (key.Length > 0)
            {
                plannedRows.TryAdd(key, row);
            }
        }

        var result = new List<NewPlanRowSpec>();
        var matches = new List<PlannedRowMatch>();

        foreach (var group in groups)
        {
            if (yesterdayLoadNumbers.Contains(group.LoadNumber))
            {
                continue;
            }

            var supplies = LoadNumberParser.ExtractSuppliesText(group.Comment);

            // Новая поставка: номера загрузки вчера не было, а в колонке «Поставки»
            // есть номер поставки. Проверяется именно текст «Поставки», то есть часть
            // комментария до слов «Номер загрузки»: то, что стоит после них, к поставке
            // отношения не имеет.
            var processing = ShipmentCodeParser.ContainsCode(supplies)
                ? OrderTextRules.CrossDockProcessing
                : string.Empty;

            // Заказ мог быть запланирован заранее отдельной строкой без номера загрузки.
            // Тогда номер вписывается в неё, а не заводится вторая такая же строка.
            var plannedKey = TextUtils.NormalizeKey(supplies);
            if (plannedKey.Length > 0 && plannedRows.Remove(plannedKey, out var plannedRow))
            {
                matches.Add(new PlannedRowMatch(plannedRow.ExcelRow, group.LoadNumber, processing));
                continue;
            }

            var section = OrderTextRules.IsReturn(group.Comment)
                ? PlanSectionKind.Returns
                : PlanSectionKind.AllGroups;

            result.Add(new NewPlanRowSpec(
                section,
                supplies,
                processing,
                OrderTextRules.LoadedComment,
                NormalizeDocumentDate(group.DocumentDate),
                group.LoadNumber));
        }

        return (result, matches);
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
        IReadOnlyList<PlannedRowMatch> plannedMatches,
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

        foreach (var match in plannedMatches)
        {
            if (seen.Add(match.LoadNumber))
            {
                loadNumbers.Add(match.LoadNumber);
            }
        }

        var updates = new List<OrderRowUpdate>(loadNumbers.Count);
        foreach (var loadNumber in loadNumbers)
        {
            // Заказ, которого сегодня нет среди обычных заказов, обнуляется по количеству
            // и помечается: строку нужно разобрать вручную.
            var inOrders = todayGroups.TryGetValue(loadNumber, out var group);
            var quantity = inOrders ? group!.Quantity : 0d;

            var outcome = JournalEvaluator.Evaluate(loadNumber, journal);
            if (!outcome.Found)
            {
                updates.Add(new OrderRowUpdate(loadNumber, 0d, null, null, !inOrders));
                continue;
            }

            var status = outcome.SetInAssembly ? OrderTextRules.InAssemblyStatus : null;
            updates.Add(new OrderRowUpdate(loadNumber, quantity, status, outcome.PercentToSet, !inOrders));
        }

        return updates;
    }

    private static IReadOnlyList<int> BuildPlaceholderDeletions(
        PlanLayout plan,
        IReadOnlyList<TodayOrderGroup> groups,
        IReadOnlyList<PlannedRowMatch> plannedMatches)
    {
        // Номера поставок берутся из той же части комментария, которая попадает
        // в колонку «Поставки»: сравниваются сопоставимые тексты.
        var realCodes = new HashSet<string>(StringComparer.Ordinal);
        foreach (var group in groups)
        {
            foreach (var code in ShipmentCodeParser.ExtractCodes(LoadNumberParser.ExtractSuppliesText(group.Comment)))
            {
                realCodes.Add(code);
            }
        }

        if (realCodes.Count == 0)
        {
            return Array.Empty<int>();
        }

        // Строку, в которую только что подставили номер загрузки, удалять нельзя:
        // заказ туда переехал, а не пришёл отдельной строкой.
        var matched = new HashSet<int>(plannedMatches.Select(m => m.ExcelRow));

        var rows = new List<int>();
        foreach (var row in plan.OrderSections.SelectMany(s => s.DataRows))
        {
            if (!row.IsOrderRow || matched.Contains(row.ExcelRow) ||
                !OrderTextRules.IsPlaceholderComment(row.Comments))
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
