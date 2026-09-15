using System.Globalization;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Правила листов с остатками и резервами: «Т.Остатки», «Остатки», «Резервы».
/// </summary>
public static class ReceivingStockRules
{
    /// <summary>
    /// Код для сравнения. На «Т.Остатках» код - текст, отрезанный от артикула формулой
    /// ПРАВСИМВ(…;6), на «Остатках» - число: «139007» и 139007 должны совпадать.
    /// </summary>
    public static string CodeKey(object? cell) => TextUtils.NormalizeKey(TextUtils.CellToString(cell));

    /// <summary>Брак на «Т.Остатках» помечен буквой «Б» в коде.</summary>
    public static bool IsDefect(object? code) =>
        CodeKey(code).Contains(ReceivingSchema.StorageStock.DefectMark);

    /// <summary>«Denny goods» = «все»: товар закрыт для всех, строка удаляется.</summary>
    public static bool IsDeniedForAll(object? value) =>
        TextUtils.EqualsKey(TextUtils.CellToString(value), ReceivingSchema.Stock.DeniedForAll);

    /// <summary>
    /// Код для «Т.Остатков», если в колонке «Код» пусто: формула не дотянута до новых
    /// строк выгрузки. Код - последние шесть знаков артикула, как в самой формуле.
    /// </summary>
    public static string CodeFromArticle(object? article)
    {
        var text = TextUtils.CellToString(article);
        return TextUtils.Normalize(text.Length <= 6 ? text : text[^6..]);
    }

    /// <summary>Сумма остатков по коду - то, что вручную делает СУММЕСЛИМН.</summary>
    public static Dictionary<string, double> SumByCode(IEnumerable<(object? Code, object? Quantity)> rows)
    {
        var sums = new Dictionary<string, double>(StringComparer.Ordinal);
        foreach (var (code, quantity) in rows)
        {
            var key = CodeKey(code);
            if (key.Length == 0 || CellError.IsError(quantity) || TextUtils.CellToDouble(quantity) is not { } value)
            {
                continue;
            }

            sums[key] = sums.TryGetValue(key, out var current) ? current + value : value;
        }

        return sums;
    }

    /// <summary>
    /// Резерв прошлых лет: «если есть данные за 2024, 2025 год, когда сейчас 2026».
    /// Дата, которая не читается, устаревшей не считается - удалять по догадке нельзя.
    /// </summary>
    public static bool IsOutdatedReserve(object? date, int currentYear) =>
        ReadYear(date) is { } year && year < currentYear;

    /// <summary>На «Резервах» должен остаться только тип «РезервОтгрузка».</summary>
    public static bool IsShipmentReserve(object? type) =>
        TextUtils.EqualsKey(
            TextUtils.CellToString(type),
            TextUtils.NormalizeKey(ReceivingSchema.Reserves.ShipmentReserve));

    private static int? ReadYear(object? value)
    {
        switch (value)
        {
            case DateTime date:
                return date.Year;
            case double number when number is > 0 and < 2958466:
                return DateTime.FromOADate(number).Year;
        }

        var text = TextUtils.Normalize(TextUtils.CellToString(value));
        var formats = new[] { "dd.MM.yyyy H:mm", "dd.MM.yyyy HH:mm", "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy", "d.M.yyyy" };
        return DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed.Year
            : null;
    }
}
