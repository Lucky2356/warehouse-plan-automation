using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Замечание осмотра книги. Помеха останавливает работу, предупреждение - нет.</summary>
public sealed record PreflightIssue(string Message, bool IsProblem);

/// <summary>
/// Осмотр книги сразу после выбора файла: всё ли на месте, прежде чем запускать
/// обработку на десятки секунд.
///
/// Проверяется ровно то, что видно снаружи, без Excel: листы и доступность связанных
/// файлов. Заголовки колонок здесь не смотрятся - для этого книгу пришлось бы открывать,
/// а это и есть та самая долгая часть.
/// </summary>
public static class PreflightCheck
{
    /// <param name="optionalSheets">
    /// Листы, без которых задача сделает не всё: название и что без него не получится.
    /// Их нехватка - предупреждение, а не помеха.
    /// </param>
    /// <param name="repeatableSheets">
    /// Обязательные листы, которых может быть несколько, и все они берутся в работу:
    /// инвойс приходит листом на каждую поставку - «Invoice-338», «Invoice-341».
    /// </param>
    public static IReadOnlyList<PreflightIssue> Run(
        WorkbookProbe probe,
        IReadOnlyList<string> requiredSheets,
        Func<string, bool> exists,
        IReadOnlyDictionary<string, string>? optionalSheets = null,
        IReadOnlyCollection<string>? repeatableSheets = null)
    {
        var issues = new List<PreflightIssue>();
        var repeatable = (repeatableSheets ?? Array.Empty<string>())
            .Select(TextUtils.NormalizeKey)
            .ToHashSet(StringComparer.Ordinal);

        foreach (var (optional, consequence) in optionalSheets ?? new Dictionary<string, string>())
        {
            if (Match(probe.Sheets, optional).Count == 0)
            {
                issues.Add(new PreflightIssue(
                    "В книге нет листа «" + optional + "»: " + consequence + ".",
                    IsProblem: false));
            }
        }

        foreach (var required in requiredSheets)
        {
            var matches = Match(probe.Sheets, required);

            if (matches.Count == 0)
            {
                issues.Add(new PreflightIssue(
                    "В книге нет листа, название которого начинается с «" + required + "».",
                    IsProblem: true));
            }
            else if (matches.Count > 1 && !repeatable.Contains(TextUtils.NormalizeKey(required)))
            {
                issues.Add(new PreflightIssue(
                    "Под «" + required + "» подходит несколько листов: " +
                    string.Join(", ", matches) + " - оставьте один.",
                    IsProblem: true));
            }
        }

        foreach (var file in probe.ExternalFiles.Where(file => !exists(file)))
        {
            issues.Add(new PreflightIssue(
                "Связанный файл сейчас недоступен: " + file +
                ". Колонки, которые из него подтягиваются, не пересчитаются.",
                IsProblem: false));
        }

        return issues;
    }

    /// <summary>
    /// Листы ищутся так же, как при обработке: точное совпадение сильнее совпадения
    /// по началу названия. Без этого «прайс» и «Прайс по подразделениям» вечно
    /// считались бы двумя кандидатами на одно и то же имя.
    /// </summary>
    private static IReadOnlyList<string> Match(IReadOnlyList<string> sheets, string required)
    {
        var key = TextUtils.NormalizeKey(required);

        var exact = sheets
            .Where(sheet => string.Equals(TextUtils.NormalizeKey(sheet), key, StringComparison.Ordinal))
            .ToList();

        return exact.Count > 0
            ? exact
            : sheets
                .Where(sheet => TextUtils.NormalizeKey(sheet).StartsWith(key, StringComparison.Ordinal))
                .ToList();
    }
}
