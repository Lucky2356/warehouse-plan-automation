using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Строка листа «на загрузку» - то, что нужно для решения по «Запрету».</summary>
/// <param name="Ban">Значение колонки «Запрет» как есть: текст, пусто или ошибка формулы.</param>
/// <param name="Available">«Фактическое кол-во, которое можно собрать на ВБ+Озон».</param>
/// <param name="Sellout">«Прогнозный sellout» в процентах: 80 - это 80 %.</param>
public sealed record RestockRow(
    int Index,
    string Sector,
    string Acr,
    double? Quantity,
    object? Ban,
    object? Available,
    double? Sellout);

/// <summary>Что программа делает со строкой.</summary>
/// <param name="Note">Что записать в «Заметка». Пусто - решения нет, оставить как есть.</param>
/// <param name="Quantity">Новое «в подтоварку». null - не менять.</param>
/// <param name="Highlight">Подсветить строку розовым.</param>
/// <param name="Reason">Почему строка попала на согласование. Пусто - не попала.</param>
public sealed record RestockDecision(
    int Index,
    string Note,
    double? Quantity,
    bool Highlight,
    string Reason,
    string? Problem)
{
    public bool NeedsApproval => Reason.Length > 0;
}

/// <summary>
/// Решение по колонке «Запрет» - самая начало разбора подтоварки.
///
/// Общее правило: можем отгрузить - в «Заметка» пишем «Ок», не можем - «Согласовать»,
/// и такая строка дополнительно попадает на лист согласования. Программа ничего
/// не додумывает: там, где данных не хватает, она оставляет «Заметку» пустой
/// и пишет замечание.
/// </summary>
public static class RestockBanRules
{
    /// <summary>Порог «Прогнозного sellout», с которого нужно согласование.</summary>
    public const double SelloutThreshold = 80d;

    public static RestockDecision Decide(RestockRow row)
    {
        var ban = TextUtils.NormalizeKey(CellError.IsError(row.Ban) ? RestockSchema.Ban.NotFound : AsText(row.Ban));

        if (ban.Length == 0)
        {
            return WithoutBan(row);
        }

        // «доп согл» стоит и само по себе, и вместе с запретом («запрет, доп согл»),
        // поэтому ищется раньше остальных значений и внутри строки, а не с её начала:
        // что бы ни стояло рядом, товар везём только после отдельного разговора.
        if (ban.Contains(RestockSchema.Ban.ExtraApproval, StringComparison.Ordinal))
        {
            return new RestockDecision(
                row.Index,
                RestockSchema.NoteApprove,
                null,
                false,
                "в «Запрете» стоит «" + AsText(row.Ban) + "»: нужно отдельное согласование",
                null);
        }

        if (ban.StartsWith(RestockSchema.Ban.OrderOnly, StringComparison.Ordinal))
        {
            return OrderOnly(row);
        }

        if (ban.StartsWith(RestockSchema.Ban.NoRetail, StringComparison.Ordinal))
        {
            return NoRetail(row);
        }

        if (ban.StartsWith(RestockSchema.Ban.NotFound, StringComparison.Ordinal))
        {
            // АЦР нет в сезонном файле - ограничений на него нет.
            return Ok(row);
        }

        return Unknown(row, "неизвестное значение «Запрет»: «" + AsText(row.Ban) + "»");
    }

    /// <summary>
    /// «Отгрузка в рамках заказа МП, запрет забора из розницы»: из розницы не берём,
    /// поэтому больше, чем можно собрать на ВБ и Озоне, не увезём. Количество урезается
    /// до собираемого, и строка подсвечивается - число в ней изменилось.
    /// </summary>
    private static RestockDecision OrderOnly(RestockRow row)
    {
        if (row.Quantity is not { } quantity || Number(row.Available) is not { } available)
        {
            return Unknown(row, "не заполнено «в подтоварку» или «Фактическое кол-во»");
        }

        if (quantity <= available)
        {
            return Ok(row);
        }

        return new RestockDecision(
            row.Index,
            RestockSchema.NoteOk,
            available,
            true,
            string.Empty,
            available > 0
                ? null
                : "собрать на ВБ+Озон нечего: «в подтоварку» обнулено, проверьте строку");
    }

    /// <summary>
    /// «Запрет забора из розницы»: не берём ничего. Бижутерию и мелкие аксессуары
    /// согласовывать не нужно, остальное идёт на лист согласования.
    /// </summary>
    private static RestockDecision NoRetail(RestockRow row)
    {
        if (IsWithoutApproval(row.Sector))
        {
            return new RestockDecision(row.Index, RestockSchema.NoteOk, null, true, string.Empty, null);
        }

        return new RestockDecision(
            row.Index,
            RestockSchema.NoteApprove,
            null,
            true,
            "запрет забора из розницы",
            null);
    }

    /// <summary>
    /// Пустой «Запрет».
    ///
    /// Разница «Фактическое кол-во» минус «в подтоварку» считается только там, где
    /// «Фактическое кол-во» заполнено: аналитик сначала фильтрует лист по непустым
    /// значениям этой колонки и только в них пишет формулу. Где собирать на ВБ+Озон
    /// нечего (пусто или ноль), разницы нет - решает «Прогнозный sellout».
    ///
    /// Проверено по готовой подтоварке: из 385 строк без запрета так разложились все.
    /// </summary>
    private static RestockDecision WithoutBan(RestockRow row)
    {
        if (row.Quantity is not { } quantity)
        {
            return Unknown(row, "не заполнено «в подтоварку»");
        }

        var available = Number(row.Available) ?? 0d;

        if (available > 0)
        {
            var difference = available - quantity;

            if (difference > 0)
            {
                return Ok(row);
            }

            if (difference < 0)
            {
                return new RestockDecision(
                    row.Index,
                    RestockSchema.NoteApprove,
                    null,
                    false,
                    "не хватает " + Format(-difference) + " шт.: можно собрать " + Format(available) +
                    ", просят " + Format(quantity),
                    null);
            }
        }

        return BySellout(row, available);
    }

    /// <summary>
    /// Решение по «Прогнозному sellout»: с 80 % и выше товар нужно согласовать,
    /// кроме бижутерии и мелких аксессуаров. Если sellout не прочитался, утверждать,
    /// что он высокий, не из чего - строка остаётся «Ок», но об этом пишется замечание.
    /// </summary>
    private static RestockDecision BySellout(RestockRow row, double available)
    {
        if (row.Sellout is not { } sellout)
        {
            return new RestockDecision(
                row.Index,
                RestockSchema.NoteOk,
                null,
                false,
                string.Empty,
                "«Прогнозный sellout» не прочитался - строка оставлена как «Ок»");
        }

        if (sellout < SelloutThreshold || IsWithoutApproval(row.Sector))
        {
            return Ok(row);
        }

        return new RestockDecision(
            row.Index,
            RestockSchema.NoteApprove,
            null,
            false,
            available > 0
                ? "остаток ровно нулевой, прогнозный sellout " + Format(sellout) + " %"
                : "собрать на ВБ+Озон нечего, прогнозный sellout " + Format(sellout) + " %",
            null);
    }

    private static RestockDecision Ok(RestockRow row) =>
        new(row.Index, RestockSchema.NoteOk, null, false, string.Empty, null);

    private static RestockDecision Unknown(RestockRow row, string problem) =>
        new(row.Index, string.Empty, null, false, string.Empty, problem);

    private static bool IsWithoutApproval(string? sector) =>
        RestockSchema.SectorsWithoutApproval.Any(known => TextUtils.EqualsKey(sector, TextUtils.NormalizeKey(known)));

    private static string AsText(object? value) => TextUtils.CellToString(value);

    /// <summary>Ошибка формулы числом не считается: «#Н/Д» приходит из COM большим числом.</summary>
    private static double? Number(object? value) =>
        CellError.IsError(value) ? null : TextUtils.CellToDouble(value);

    private static string Format(double value) =>
        value.ToString("0.##", System.Globalization.CultureInfo.CurrentCulture);
}
