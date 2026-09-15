using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Что программа прочитала из колонок, которые заполняет аналитик по стенкам.</summary>
/// <param name="StorePeriod">Значение «Период продаж магазины» как есть.</param>
/// <param name="IslandPeriod">Значение «Период продаж острова» как есть.</param>
/// <param name="WallAnalogue">«Аналог стенки».</param>
/// <param name="GoogleAnalogue">«Аналог гугл».</param>
/// <param name="PriceCheck">«Проверка цен»: ИСТИНА, если согласованная цена сошлась с розничной.</param>
/// <param name="Grades">Грейды, у которых в строке стоит 1.</param>
/// <param name="Marketplace">«МП» из стенок. null - колонки на листе нет.</param>
/// <param name="GradeColumns">
/// Грейды, колонки которых есть на листе «Цены». null - все, что знает программа.
/// Грейд без колонки не сверяется: нельзя сказать, что он «не отмечен».
/// </param>
public sealed record PriceRowState(
    int Index,
    string Acr,
    object? StorePeriod,
    object? IslandPeriod,
    object? WallAnalogue,
    object? GoogleAnalogue,
    object? PriceCheck,
    IReadOnlyList<string> Grades,
    object? Marketplace = null,
    IReadOnlyList<string>? GradeColumns = null);

/// <summary>Итог проверок одной строки.</summary>
public sealed record PriceRowCheck(
    int Index,
    string LinkValue,
    string BanCheck,
    IReadOnlyList<string> Highlight,
    IReadOnlyList<ProcessingWarning> Warnings);

/// <summary>
/// Проверки листа «Цены».
///
/// Колонки со стенками заполняет аналитик, а программа их читает, сверяет между собой
/// и с листом «link» и подсвечивает то, что не сошлось. Ничего не додумывается: там,
/// где данных нет, пишется замечание.
/// </summary>
public static class PriceChecks
{
    public const string LinkPresent = "есть";
    public const string LinkMissing = "нет";

    /// <summary>
    /// Проверки этапа подготовки: только то, что видно без стенок. Колонки со стенками
    /// аналитик протягивает после него, и судить по ним пока не о чем.
    /// Наценок здесь тоже ещё нет - их ставит пересчёт.
    /// Замечания здесь не заводятся: про ненайденный штрихкод и повтор уже написал
    /// сборщик строк, и повторять его незачем.
    /// </summary>
    public static PriceRowCheck Prepare(PriceRowValues values)
    {
        var highlight = new List<string>();

        if (values.Reference is null)
        {
            highlight.Add(PriceSchema.Prices.BarcodeCopy);
        }

        if (values.IsDuplicateBarcode)
        {
            highlight.Add(PriceSchema.Prices.PurchasePrice);
        }

        return new PriceRowCheck(
            values.Index, string.Empty, string.Empty, highlight, Array.Empty<ProcessingWarning>());
    }

    public static PriceRowCheck Check(
        PriceRowValues values,
        PriceRowState state,
        LinkEntry? link)
    {
        var highlight = new List<string>();
        var warnings = new List<ProcessingWarning>();
        var notes = new List<string>();
        var where = PriceSchema.PricesSheet + ", штрихкод " + values.Barcode;

        if (values.Reference is null)
        {
            highlight.Add(PriceSchema.Prices.BarcodeCopy);
        }

        if (values.IsDuplicateBarcode)
        {
            highlight.Add(PriceSchema.Prices.PurchasePrice);
        }

        if (values.Markup is null)
        {
            highlight.Add(PriceSchema.Prices.PlannedMarkup);
        }

        CheckSalesPeriod(state, highlight, warnings, where);
        CheckAnalogues(state, highlight, warnings, where);
        CheckPriceMatch(state, highlight, warnings, where);
        CheckMarketplace(values, state, highlight, warnings, where);

        // Запреты отдельных магазинов - обычное дело, они пишутся в «Проверку запретов»,
        // но в замечания не выносятся: иначе замечаний было бы по строке на каждый АЦР.
        var storeNotes = new List<string>();
        var linkValue = link is null ? LinkMissing : LinkPresent;
        var ok = false;

        if (link is null)
        {
            highlight.Add(PriceSchema.Prices.Link);
            notes.Add("Нужно добавить в Линк.");
        }
        else
        {
            var datesMatch = CheckLinkDates(state, link, notes, highlight);
            CheckDenials(state, link, notes, storeNotes);
            CheckGrades(state, link, notes, highlight);

            // «Ок» - когда даты совпали со стенками, а в «link» АЦР включён и не запрещён везде,
            // где он положен. Грейд, который в «Цены» не отмечен (O120 и O140 = 0), в «link»
            // и должен быть выключен: его запрет «Ок» не мешает.
            ok = datesMatch && notes.Count == 0 && storeNotes.Count == 0;
        }

        foreach (var note in notes)
        {
            warnings.Add(new ProcessingWarning(note, where));
        }

        var banCheck = ok ? BanCheckOk : string.Join(" ", notes.Concat(storeNotes));

        return new PriceRowCheck(
            values.Index,
            linkValue,
            banCheck,
            highlight,
            warnings);
    }

    /// <summary>Что пишется в «Проверку запретов», когда в «link» всё сошлось.</summary>
    public const string BanCheckOk = "Ок";

    /// <summary>
    /// Запреты по грейдам магазинов. Грейд, у которого АЦР запрещён во всех магазинах, - это
    /// запрет грейда, о нём замечание. Запрет в части магазинов грейда пишется в колонку.
    ///
    /// Запрет грейда, который в «Цены» не отмечен, - не расхождение, а то, что и должно быть:
    /// у АЦР, который не идёт на острова, O120 и O140 = 0, и в «link» острова выключены.
    /// Об этом ничего не пишется. Грейд без колонки на листе не отмечен ни так, ни так -
    /// о его запрете по-прежнему замечание.
    /// </summary>
    private static void CheckDenials(PriceRowState state, LinkEntry link, List<string> notes, List<string> storeNotes)
    {
        var columns = (state.GradeColumns ?? PriceSchema.Prices.AllGrades)
            .Select(LinkSheetReader.GradeKey)
            .ToHashSet(StringComparer.Ordinal);
        var marked = state.Grades.Select(LinkSheetReader.GradeKey).ToHashSet(StringComparer.Ordinal);
        bool Expected(LinkGrade grade)
        {
            // Острова без грейда в «link» - это O120 и O140 вместе: они положены, если в «Цены»
            // отмечен хоть один из них (или колонок островов на листе нет - тогда не проверить).
            if (grade.Grade == PriceSchema.Link.IslandsWithoutGrade)
            {
                var islands = PriceSchema.Prices.IslandGrades.Select(LinkSheetReader.GradeKey).ToList();
                return !islands.Any(columns.Contains) || islands.Any(marked.Contains);
            }

            var key = LinkSheetReader.GradeKey(grade.Grade);
            return !columns.Contains(key) || marked.Contains(key);
        }

        if (link.Rows > 0 && link.AllowedRows == 0)
        {
            notes.Add("Во всех магазинах «link» этот АЦР запрещён (Вкл = 0 либо Deny goods = 1).");
            return;
        }

        // Магазины без «Группа_ам» грейдом не являются: о них - только сводка в колонке.
        var ungraded = link.Grades.FirstOrDefault(grade => grade.Grade.Length == 0);
        if (ungraded is not null && ungraded.AllowedStores < ungraded.Stores)
        {
            storeNotes.Add(
                "Запрет в магазинах без «Группа_ам»: " + (ungraded.Stores - ungraded.AllowedStores) +
                " из " + ungraded.Stores + ".");
        }

        var graded = link.Grades.Where(grade => grade.Grade.Length > 0).ToList();
        var denied = graded.Where(grade => grade.IsFullyDenied && Expected(grade)).ToList();
        if (denied.Count > 0)
        {
            notes.Add(
                "Запрет на грейд" + (denied.Count == 1 ? "е " : "ах ") +
                string.Join(", ", denied.Select(grade => grade.Grade)) +
                ": в «link» АЦР запрещён во всех магазинах " + (denied.Count == 1 ? "этого грейда" : "этих грейдов") + ".");
        }

        // Острова, которые в «Цены» не отмечены, АЦР не положены - частичный запрет по ним тоже не новость.
        var partly = graded
            .Where(grade => grade.IsPartlyDenied && (grade.Grade != PriceSchema.Link.IslandsWithoutGrade || Expected(grade)))
            .ToList();
        if (partly.Count > 0)
        {
            storeNotes.Add(
                "Запрет в части магазинов (запрещено из всех): " +
                string.Join(", ", partly.Select(grade =>
                    grade.Grade + " - " + (grade.Stores - grade.AllowedStores) + " из " + grade.Stores)) +
                ".");
        }
    }

    private static void CheckSalesPeriod(
        PriceRowState state,
        List<string> highlight,
        List<ProcessingWarning> warnings,
        string where)
    {
        // Разорванная ссылка - это не «позиции нет в стенках», а «файл стенок не прочитался».
        // Он лежит на сетевом диске, и без доступа к нему Excel теряет сохранённые значения
        // при первом же пересчёте. Проверяется и колонка островов: ссылки на стенки в ней те же.
        if (CellError.IsBrokenReference(state.StorePeriod) || CellError.IsBrokenReference(state.IslandPeriod))
        {
            highlight.Add(PriceSchema.Prices.StoreSalesPeriod);
            warnings.Add(new ProcessingWarning(
                "Колонки со стенками не прочитались: «#ССЫЛКА!» вместо значений. Обычно это " +
                "значит, что сетевая папка со стенками была недоступна. Проверьте доступ " +
                "и протяните эти колонки заново.",
                where));
            return;
        }

        if (CellError.IsError(state.StorePeriod))
        {
            highlight.Add(PriceSchema.Prices.StoreSalesPeriod);
            warnings.Add(new ProcessingWarning(
                "«Период продаж магазины» не подтянулся из стенок: этой позиции в стенках нет " +
                "либо она записана там с ошибкой.",
                where));
            return;
        }

        var text = TextUtils.Normalize(TextUtils.CellToString(state.StorePeriod));
        if (text.Length == 0 || text == "0")
        {
            highlight.Add(PriceSchema.Prices.StoreSalesPeriod);
            warnings.Add(new ProcessingWarning(
                "«Период продаж магазины» пуст или равен нулю: проверьте позицию в стенках вручную.",
                where));
        }
    }

    private static void CheckAnalogues(
        PriceRowState state,
        List<string> highlight,
        List<ProcessingWarning> warnings,
        string where)
    {
        var wall = Analogue(state.WallAnalogue);
        var google = Analogue(state.GoogleAnalogue);

        if (wall == google)
        {
            return;
        }

        highlight.Add(PriceSchema.Prices.GoogleAnalogue);
        warnings.Add(new ProcessingWarning(
            "Аналоги не сходятся: в стенках «" + Describe(state.WallAnalogue) +
            "», в гугле «" + Describe(state.GoogleAnalogue) + "».",
            where));
    }

    /// <summary>
    /// В обеих колонках аналогов смысл двоичный: либо аналога нет («нет», 0, пусто, «#Н/Д»),
    /// либо он есть и записан своим названием. Сравниваются именно эти два состояния.
    /// </summary>
    private static bool Analogue(object? value)
    {
        if (value is null || CellError.IsError(value))
        {
            return false;
        }

        var text = TextUtils.NormalizeKey(TextUtils.CellToString(value));
        return text.Length > 0 && text != "0" && text != "нет";
    }

    private static string Describe(object? value) =>
        CellError.IsError(value) ? "#Н/Д" : TextUtils.Normalize(TextUtils.CellToString(value));

    private static void CheckPriceMatch(
        PriceRowState state,
        List<string> highlight,
        List<ProcessingWarning> warnings,
        string where)
    {
        if (state.PriceCheck is null || CellError.IsError(state.PriceCheck))
        {
            highlight.Add(PriceSchema.Prices.PriceCheck);
            warnings.Add(new ProcessingWarning(
                "«Проверка цен» не посчиталась: скорее всего не подтянулась согласованная цена.",
                where));
            return;
        }

        if (IsFalse(state.PriceCheck))
        {
            highlight.Add(PriceSchema.Prices.PriceCheck);
            warnings.Add(new ProcessingWarning(
                "Согласованная цена не совпадает с розничной ценой сети.",
                where));
        }
    }

    /// <summary>
    /// «МП» равно количеству поставки - возможно, вся поставка идёт на маркетплейс.
    /// Узнать это можно только в слежении или в «Отчёте по заказам» МЦ, поэтому формулу
    /// «Кол-во за вычетом мп и блогеров» программа не меняет, а показывает строку.
    /// </summary>
    private static void CheckMarketplace(
        PriceRowValues values,
        PriceRowState state,
        List<string> highlight,
        List<ProcessingWarning> warnings,
        string where)
    {
        if (state.Marketplace is null || CellError.IsError(state.Marketplace) || values.Units <= 0)
        {
            return;
        }

        var marketplace = TextUtils.CellToDouble(state.Marketplace);
        if (marketplace is null || Math.Abs(marketplace.Value - values.Units) > 1e-9)
        {
            return;
        }

        highlight.Add(PriceSchema.Prices.Marketplace);
        warnings.Add(new ProcessingWarning(
            "«МП» равно количеству поставки (" + TextUtils.CellToString(values.Units) + " шт.): " +
            "возможно, это поставка МП. Проверьте в слежении или в «Отчёте по заказам» МЦ; " +
            "если да - в «Кол-во за вычетом мп и блогеров» формула меняется на «Ед. минус МП».",
            where));
    }

    private static bool IsFalse(object? value) => value switch
    {
        bool flag => !flag,
        double number => Math.Abs(number) < 1e-9,
        string text => TextUtils.EqualsKey(text, "ложь") || TextUtils.EqualsKey(text, "false"),
        _ => false,
    };

    /// <summary>Сверяет даты «link» с «Период продаж магазины». true - даты есть и совпали.</summary>
    private static bool CheckLinkDates(
        PriceRowState state,
        LinkEntry link,
        List<string> notes,
        List<string> highlight)
    {
        if (link.DateStart is null || link.DateEnd is null)
        {
            return false;
        }

        if (!SalesPeriodParser.TryParse(TextUtils.CellToString(state.StorePeriod), out var period))
        {
            return false;
        }

        var start = DateTime.FromOADate(link.DateStart.Value);
        var end = DateTime.FromOADate(link.DateEnd.Value);
        if (SalesPeriodParser.Matches(period, start, end))
        {
            return true;
        }

        highlight.Add(PriceSchema.Prices.StoreSalesPeriod);
        notes.Add(
            "Даты не совпадают: в «link» " + SalesPeriodParser.FromDates(start, end) +
            ", в «Цены» " + period + ".");
        return false;
    }

    /// <summary>
    /// Грейды «Цены» против грейдов магазинов «link». В «Цены» 1 в колонке грейда значит
    /// «этому грейду АЦР положен»; в «link» грейд положен, если хотя бы в одном его магазине
    /// АЦР включён и не запрещён. Сверяются только грейды, колонки которых есть на листе.
    /// </summary>
    private static void CheckGrades(
        PriceRowState state,
        LinkEntry link,
        List<string> notes,
        List<string> highlight)
    {
        var columns = (state.GradeColumns ?? PriceSchema.Prices.AllGrades)
            .ToDictionary(LinkSheetReader.GradeKey, grade => grade, StringComparer.Ordinal);
        var marked = state.Grades.Select(LinkSheetReader.GradeKey).ToHashSet(StringComparer.Ordinal);
        var inLink = link.Grades
            .Where(grade => grade.Grade.Length > 0)
            .ToDictionary(grade => LinkSheetReader.GradeKey(grade.Grade), StringComparer.Ordinal);

        var notMarked = new List<string>();
        var notInLink = new List<string>();

        foreach (var (key, column) in columns)
        {
            inLink.TryGetValue(key, out var grade);
            var isMarked = marked.Contains(key);

            if (isMarked && grade is null)
            {
                notInLink.Add(column);
                highlight.Add(column);
            }
            else if (isMarked && grade!.IsFullyDenied)
            {
                // О самом запрете уже написано - здесь только пометка на колонке грейда.
                highlight.Add(column);
            }
            else if (!isMarked && grade is { IsFullyDenied: false })
            {
                notMarked.Add(column);
            }
        }

        if (notInLink.Count > 0)
        {
            notes.Add(
                "Отмечены в «Цены», но в «link» магазинов этих грейдов нет: " + string.Join(", ", notInLink) + ".");
        }

        if (notMarked.Count > 0)
        {
            notes.Add(
                "В «link» положены, но в «Цены» не отмечены: " + string.Join(", ", notMarked) + ".");
        }
    }
}
