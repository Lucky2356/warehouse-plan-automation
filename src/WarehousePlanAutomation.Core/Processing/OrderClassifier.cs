using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Классификация строк листа «Заказы на отгрузку» строго в порядке, заданном ТЗ:
/// «Опт» - «777» - маркетплейсы (возвраты, из поставок, хранение) - документы «ЗП» -
/// служебные строки - реальные заказы по номеру загрузки.
/// Всё, что не подошло, остаётся на листе.
/// </summary>
public static class OrderClassifier
{
    public static OrderClassificationResult Classify(IReadOnlyList<OrderRow> rows)
    {
        var totals = new Dictionary<OrderCategory, double>();
        var rowsToDelete = new List<int>();
        var leftovers = new List<OrderRow>();
        var realOrders = new List<(OrderRow Row, long LoadNumber)>();

        foreach (var row in rows)
        {
            var category = Categorize(row, out var loadNumber);
            Accumulate(totals, category, row.DifferenceUnits);

            switch (category)
            {
                case OrderCategory.Unresolved:
                    leftovers.Add(row);
                    break;
                case OrderCategory.RealOrder:
                    realOrders.Add((row, loadNumber));
                    rowsToDelete.Add(row.ExcelRow);
                    break;
                default:
                    rowsToDelete.Add(row.ExcelRow);
                    break;
            }
        }

        return new OrderClassificationResult(totals, rowsToDelete, GroupOrders(realOrders), leftovers);
    }

    public static OrderCategory Categorize(OrderRow row, out long loadNumber)
    {
        loadNumber = 0;

        if (OrderTextRules.IsWholesale(row.Division))
        {
            return OrderCategory.Wholesale;
        }

        if (OrderTextRules.IsInternetShop(row.Division))
        {
            return OrderCategory.InternetShop;
        }

        if (OrderTextRules.IsMarketplace(row.Division))
        {
            if (OrderTextRules.IsReturn(row.Comment))
            {
                return OrderCategory.MarketplaceReturns;
            }

            return ShipmentCodeParser.ContainsCode(row.Comment)
                ? OrderCategory.MarketplaceSupplies
                : OrderCategory.MarketplaceStorage;
        }

        // Проверка идёт после подразделений: строки «Опт», «777» и маркетплейсов
        // и без того удаляются, но их количества попадают в итоги блоков, и признак
        // «ЗП» в комментарии не повод выкидывать их из этих сумм.
        if (OrderTextRules.IsZpRow(row.Comment))
        {
            return OrderCategory.ZpDocument;
        }

        if (OrderTextRules.IsServiceRow(row.Comment))
        {
            return OrderCategory.Service;
        }

        return LoadNumberParser.TryExtract(row.Comment, out loadNumber)
            ? OrderCategory.RealOrder
            : OrderCategory.Unresolved;
    }

    private static void Accumulate(Dictionary<OrderCategory, double> totals, OrderCategory category, double value)
    {
        totals[category] = totals.TryGetValue(category, out var current) ? current + value : value;
    }

    private static IReadOnlyList<TodayOrderGroup> GroupOrders(IReadOnlyList<(OrderRow Row, long LoadNumber)> orders)
    {
        var order = new List<long>();
        var quantities = new Dictionary<long, double>();
        var firstRows = new Dictionary<long, OrderRow>();
        var sourceRows = new Dictionary<long, List<int>>();

        foreach (var (row, loadNumber) in orders)
        {
            if (!quantities.ContainsKey(loadNumber))
            {
                order.Add(loadNumber);
                quantities[loadNumber] = 0d;
                firstRows[loadNumber] = row;
                sourceRows[loadNumber] = new List<int>();
            }

            quantities[loadNumber] += row.DifferenceUnits;
            sourceRows[loadNumber].Add(row.ExcelRow);
        }

        return order
            .Select(loadNumber => new TodayOrderGroup(
                loadNumber,
                quantities[loadNumber],
                firstRows[loadNumber].Comment,
                firstRows[loadNumber].DocumentDate,
                sourceRows[loadNumber]))
            .ToList();
    }
}
