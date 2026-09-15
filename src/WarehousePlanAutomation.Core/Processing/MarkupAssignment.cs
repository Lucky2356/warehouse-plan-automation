using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Наценки для строк листа «Цены» и замечания по тем, кому наценка не досталась.</summary>
public sealed record MarkupPlan(
    IReadOnlyList<PriceRowValues> Rows,
    IReadOnlyList<ProcessingWarning> Warnings);

/// <summary>
/// Расстановка наценок по строкам листа «Цены».
///
/// Наценки ставятся на втором этапе распреда: к этому времени подгруппы уже проставлены
/// и видно, сколько их в поставке. Наценка берётся на каждую подгруппу свою - в одной
/// поставке их может быть несколько, и регламент для них разный.
///
/// Там, где регламент ответа не даёт, строка остаётся такой, какой пришла с листа:
/// наценку в ней могла проставить рука человека, и стирать её нельзя.
/// </summary>
public static class MarkupAssignment
{
    public static MarkupPlan Apply(IReadOnlyList<PriceRowValues> rows, MarkupResolver resolver)
    {
        var warnings = new List<ProcessingWarning>();
        var reported = new HashSet<string>(StringComparer.Ordinal);
        var result = new List<PriceRowValues>(rows.Count);

        foreach (var row in rows)
        {
            var choice = resolver.Resolve(row.Reference?.Sector, row.Subgroup);

            if (choice.Warning is not null && reported.Add(choice.Warning))
            {
                warnings.Add(new ProcessingWarning(choice.Warning, PriceSchema.MarkupSheet));
            }

            // Наценку, которой в регламенте нет, аналитик проставляет руками - и то,
            // что она вписала, остаётся: пересчёт запускают не один раз, и на втором
            // запуске её цифры пропали бы молча.
            result.Add(choice.Rule is null ? row : row with { Markup = choice.Rule });
        }

        return new MarkupPlan(result, warnings);
    }
}
