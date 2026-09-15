using System.Globalization;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>Строка листа «КС»: доли продаж по неделям для одной подгруппы.</summary>
public sealed record SeasonalityRow(
    string Sector,
    string Group,
    string Subgroup,
    IReadOnlyDictionary<int, double> Weeks);

/// <summary>
/// Чтение листа «КС». Номера недель стоят прямо в строке заголовков, последняя колонка -
/// итог, а не неделя, поэтому нечисловые заголовки пропускаются.
/// </summary>
public static class SeasonalitySheetReader
{
    public static IReadOnlyList<SeasonalityRow> Read(SheetGrid grid)
    {
        var headers = HeaderResolver.Resolve(grid, PriceSchema.SeasonalitySheet, PriceSchema.Seasonality.Specs);

        var sectorColumn = headers[PriceSchema.Seasonality.Sector];
        var groupColumn = headers[PriceSchema.Seasonality.Group];
        var subgroupColumn = headers[PriceSchema.Seasonality.Subgroup];

        var weekColumns = ReadWeekColumns(grid, headers.HeaderRow, subgroupColumn);

        var rows = new List<SeasonalityRow>();
        for (var row = headers.HeaderRow + 1; row <= grid.LastRow; row++)
        {
            var subgroup = TextUtils.Normalize(grid.Text(row, subgroupColumn));
            if (subgroup.Length == 0)
            {
                continue;
            }

            var weeks = new Dictionary<int, double>();
            foreach (var pair in weekColumns)
            {
                var value = grid.Number(row, pair.Value);
                if (value is not null)
                {
                    weeks[pair.Key] = value.Value;
                }
            }

            rows.Add(new SeasonalityRow(
                TextUtils.Normalize(grid.Text(row, sectorColumn)),
                TextUtils.Normalize(grid.Text(row, groupColumn)),
                subgroup,
                weeks));
        }

        return rows;
    }

    /// <summary>Колонки недель: заголовок - целое число от 1 до 53.</summary>
    private static Dictionary<int, int> ReadWeekColumns(SheetGrid grid, int headerRow, int afterColumn)
    {
        var result = new Dictionary<int, int>();

        for (var column = afterColumn + 1; column <= grid.LastColumn; column++)
        {
            var text = TextUtils.Normalize(grid.Text(headerRow, column));
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var week) &&
                week is >= 1 and <= 53)
            {
                result[week] = column;
            }
        }

        return result;
    }
}
