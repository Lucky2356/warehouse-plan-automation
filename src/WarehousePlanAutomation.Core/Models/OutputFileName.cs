using System.Globalization;
using System.Text.RegularExpressions;

namespace WarehousePlanAutomation.Core.Models;

/// <summary>
/// Имя файла результата: к имени книги добавляется, что с ней сделали, и когда.
///
/// Отметка прошлого запуска при этом снимается. Распред идёт двумя этапами, и файл
/// второго делается из файла первого: без этого имена накапливались бы -
/// «Поставка подготовка 2026-09-09_1550 распред 2026-09-09_1552».
/// </summary>
public static class OutputFileName
{
    /// <summary>Что сделали с книгой - это слово попадает в имя файла.</summary>
    public const string PlanMark = "план склада";

    public const string PrepareMark = "подготовка";

    public const string DistributionMark = "распред";

    public const string RestockMark = "подтоварка";

    /// <summary>Первый этап приемки на хранилище.</summary>
    public const string ReceivingMark = "приемка";

    /// <summary>Второй этап приемки: адреса и тары.</summary>
    public const string AddressesMark = "адреса";

    /// <summary>
    /// Отметка, оставленная программой: слово, дата и время. «готово» - от версий
    /// до разделения распреда на этапы: такие файлы тоже должны переименовываться начисто.
    /// </summary>
    private static readonly Regex Mark = new(
        @"[ _](готово|план склада|подготовка|распред|подтоварка|приемка|адреса)[ _]\d{4}-\d{2}-\d{2}[ _]\d{4,6}(_\d+)?$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>Имя книги без отметки предыдущего запуска.</summary>
    public static string WithoutMark(string fileName) =>
        Mark.Replace(fileName ?? string.Empty, string.Empty).TrimEnd();

    /// <summary>Имя файла результата без расширения.</summary>
    public static string Build(string fileName, string mark, DateTime now) =>
        WithoutMark(fileName) + " " + mark + " " +
        now.ToString("yyyy-MM-dd_HHmm", CultureInfo.InvariantCulture);
}
