using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Что делается со строкой «Приходов».</summary>
public enum IncomingKind
{
    /// <summary>Строка удаляется: статус не «Запущено», разница мала или это поставка «Л».</summary>
    Removed,

    /// <summary>Поставка маркетплейса - розовая, идёт на «Поставки МП».</summary>
    Marketplace,

    /// <summary>Поставка уже собрана - зелёная, идёт на «Поставки собраны».</summary>
    Collected,

    /// <summary>Поставка ещё в плане склада - голубая, идёт на «Поставки не собраны».</summary>
    NotCollected,

    /// <summary>Разобрать не получилось: строка остаётся без цвета, об этом пишется замечание.</summary>
    Unresolved,
}

/// <summary>Строка листа «Приходы» - то, что нужно для решения.</summary>
public sealed record IncomingRow(int Index, string Number, object? Status, object? Difference);

/// <summary>Решение по строке. <paramref name="Note"/> - то, что стоит сказать человеку.</summary>
public sealed record IncomingDecision(int Index, IncomingKind Kind, string Note);

/// <summary>
/// Разбор листа «Приходы».
///
/// Сначала лишнее удаляется: всё, что не «Запущено», разница до 9 единиц включительно
/// и поставки на «Л». Остальное красится: «М» - маркетплейс; «С», найденная в «Удалили
/// из плана склада», - собрана; «С», найденная в «Плане склада», - не собрана.
/// Поиск в плане идёт вторым и перекрашивает первый - так же, как при ручном разборе.
/// </summary>
public static class IncomingSupplyRules
{
    /// <summary>Разница, начиная с которой строка остаётся: «до 9 включительно» удаляются.</summary>
    public const double MinimumDifference = 10d;

    private const string RunningStatus = "запущен";

    public static IReadOnlyList<IncomingDecision> Decide(
        IReadOnlyList<IncomingRow> rows,
        IReadOnlySet<string> removedFromPlan,
        IReadOnlySet<string> inPlan)
    {
        var decisions = new List<IncomingDecision>(rows.Count);

        foreach (var row in rows)
        {
            decisions.Add(DecideOne(row, removedFromPlan, inPlan));
        }

        return decisions;
    }

    private static IncomingDecision DecideOne(
        IncomingRow row,
        IReadOnlySet<string> removedFromPlan,
        IReadOnlySet<string> inPlan)
    {
        var status = TextUtils.Normalize(TextUtils.CellToString(row.Status));
        if (!TextUtils.NormalizeKey(status).StartsWith(RunningStatus, StringComparison.Ordinal))
        {
            return new IncomingDecision(row.Index, IncomingKind.Removed, "статус «" + status + "»");
        }

        var difference = CellError.IsError(row.Difference) ? null : TextUtils.CellToDouble(row.Difference);
        if (difference is { } value && value < MinimumDifference)
        {
            return new IncomingDecision(
                row.Index, IncomingKind.Removed, "разница " + value.ToString("0.##") + " ед.");
        }

        var letter = SupplyNumbers.Letter(row.Number);
        if (letter == SupplyNumbers.Excluded)
        {
            return new IncomingDecision(row.Index, IncomingKind.Removed, "поставка на «Л»");
        }

        if (letter == SupplyNumbers.Marketplace)
        {
            return new IncomingDecision(row.Index, IncomingKind.Marketplace, string.Empty);
        }

        var digits = SupplyNumbers.Digits(row.Number);
        if (letter != SupplyNumbers.Network || digits.Length == 0)
        {
            return new IncomingDecision(
                row.Index,
                IncomingKind.Unresolved,
                "номер «" + row.Number + "» не начинается с «С», «М» или «Л»");
        }

        var removed = removedFromPlan.Contains(digits);
        if (inPlan.Contains(digits))
        {
            return new IncomingDecision(
                row.Index,
                IncomingKind.NotCollected,
                removed
                    ? "поставка " + digits + " есть и в «Удалили из плана склада», и в «Плане склада» - " +
                      "считается не собранной: план проверяется вторым"
                    : string.Empty);
        }

        return removed
            ? new IncomingDecision(row.Index, IncomingKind.Collected, string.Empty)
            : new IncomingDecision(
                row.Index,
                IncomingKind.Unresolved,
                "поставки " + digits + " нет ни в «Удалили из плана склада», ни в «Плане склада»");
    }
}
