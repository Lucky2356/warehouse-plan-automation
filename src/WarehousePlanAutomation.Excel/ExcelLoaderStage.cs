using System.Globalization;
using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Logging;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>Что получилось на третьем этапе.</summary>
internal sealed record LoaderOutcome(
    int Columns,
    bool? Check,
    IReadOnlyList<ProcessingWarning> Warnings);

/// <summary>
/// Третий этап распреда: лист «Загрузочник».
///
/// Работы здесь немного, потому что почти всё делают сами формулы листа. Нужно только
/// довести число колонок до числа кодов, размножить по ним первую колонку и вписать
/// коды в строку заголовков. Формулы «Загрузочника» смотрят на лист «Распред» - и это
/// важно для порядка: колонки там уже переставлены вторым этапом, поэтому ссылки вида
/// Распред!$N$28:$Y$147 Excel растянул сам, вручную их править не нужно.
/// </summary>
internal sealed class ExcelLoaderStage
{
    private readonly IAppLogger _logger;

    public ExcelLoaderStage(IAppLogger logger)
    {
        _logger = logger;
    }

    public LoaderOutcome Run(
        object application,
        object sheet,
        IReadOnlyList<PriceRowValues> priceRows)
    {
        var warnings = new List<ProcessingWarning>();
        var layout = ReadLayout(sheet);

        if (priceRows.Count == 0)
        {
            warnings.Add(new ProcessingWarning(
                "В поставке нет ни одного кода: колонки на листе «" + PriceSchema.LoaderSheet +
                "» не перестроены.",
                PriceSchema.LoaderSheet));
            return new LoaderOutcome(layout.CodeColumnCount, null, warnings);
        }

        var before = layout.CodeColumnCount;
        layout = Resize(sheet, layout, priceRows.Count);
        RepairCollapsedSums(sheet, layout, before);
        CopyFirstColumn(application, sheet, layout);
        WriteCodes(sheet, layout, priceRows, warnings);

        return new LoaderOutcome(layout.CodeColumnCount, null, warnings);
    }

    /// <summary>
    /// Читает ячейку проверки. Вызывается последним - после того как посчитаны и распред,
    /// и «мин запас на Хаб»: до пересчёта в ячейке стоит значение прошлой поставки.
    /// </summary>
    public LoaderOutcome Verify(object sheet, LoaderOutcome outcome)
    {
        var layout = ReadLayout(sheet);
        var raw = ExcelSheetOperations.GetValue(sheet, layout.CheckCell.Row, layout.CheckCell.Column);
        var warnings = outcome.Warnings.ToList();

        if (raw is bool value)
        {
            if (!value)
            {
                warnings.Add(new ProcessingWarning(
                    "На листе «" + PriceSchema.LoaderSheet + "» проверка дала «ЛОЖЬ»: суммы " +
                    "загрузочника и листа «" + PriceSchema.DistributionSheet + "» не сошлись. " +
                    "Загрузочник отдавать нельзя, пока там не появится «ИСТИНА».",
                    PriceSchema.LoaderSheet + ", " + layout.CheckCell,
                    ExcelSheetOperations.GetSheetName(sheet),
                    layout.CheckCell.ToString()));
            }

            _logger.Information(
                "Проверка «" + PriceSchema.LoaderSheet + "» в " + layout.CheckCell + ": " +
                (value ? "ИСТИНА" : "ЛОЖЬ") + ".");

            return outcome with { Check = value, Warnings = warnings };
        }

        warnings.Add(new ProcessingWarning(
            "На листе «" + PriceSchema.LoaderSheet + "» проверка не посчиталась: в ячейке " +
            (CellError.IsError(raw) ? "ошибка формулы" : "не «ИСТИНА» и не «ЛОЖЬ»") +
            ". Проверьте загрузочник вручную.",
            PriceSchema.LoaderSheet + ", " + layout.CheckCell,
            ExcelSheetOperations.GetSheetName(sheet),
            layout.CheckCell.ToString()));

        return outcome with { Check = null, Warnings = warnings };
    }

    private static LoaderLayout ReadLayout(object sheet) =>
        LoaderLayout.Read(ExcelSheetOperations.ReadGrid(sheet, withFormulas: true));

    /// <summary>
    /// Доводит число колонок до числа кодов. Как и на листе «Распред», колонки вставляются
    /// и удаляются внутри уже занятого диапазона: тогда Excel сам растягивает формулы,
    /// которые охватывают все колонки сразу - «Итого» в строках РТТ и суммы над таблицей.
    /// </summary>
    private LoaderLayout Resize(object sheet, LoaderLayout layout, int wanted)
    {
        var difference = wanted - layout.CodeColumnCount;
        if (difference == 0)
        {
            return layout;
        }

        if (difference > 0)
        {
            // Обычно колонки вставляются внутрь занятого диапазона - тогда Excel сам
            // растягивает суммы по всем колонкам. Но когда колонка кодов одна, «внутри»
            // не существует: вставка в неё саму сдвинула бы её вправо вместе с данными.
            // Тогда вставляем справа от неё, а суммы чинит RepairCollapsedSums.
            var insertAt = layout.CodeColumnCount == 1 ? layout.TotalColumn : layout.LastCodeColumn;

            ExcelSheetOperations.InsertColumns(sheet, insertAt, difference);
            _logger.Information(
                "На листе «" + PriceSchema.LoaderSheet + "» добавлено колонок: " + difference + ".");
        }
        else
        {
            ExcelSheetOperations.DeleteColumns(sheet, layout.CodeColumnAt(wanted), -difference);
            _logger.Information(
                "На листе «" + PriceSchema.LoaderSheet + "» удалено колонок: " + -difference + ".");
        }

        // Разметка не перечитывается: у пустых колонок ещё нет кодов, и по листу их
        // не сосчитать. Сколько их должно быть, знает только этот метод.
        return layout.WithCodeColumnCount(wanted);
    }

    /// <summary>
    /// Чинит суммы, которые Excel не смог растянуть сам.
    ///
    /// Пока колонок кодов было больше одной, вставка внутрь диапазона расширяет его.
    /// Но диапазон из одной ячейки («СУММ(G13:G13)») расширить нельзя - он просто
    /// уезжает вправо и продолжает считать одну колонку. Так бывает после поставки
    /// с единственным кодом: её загрузочник становится заготовкой для следующей.
    /// </summary>
    private void RepairCollapsedSums(object sheet, LoaderLayout layout, int columnsBefore)
    {
        if (columnsBefore != 1 || layout.CodeColumnCount <= 1)
        {
            return;
        }

        var grid = ExcelSheetOperations.ReadGrid(sheet, withFormulas: true);
        var repaired = 0;

        for (var row = grid.FirstRow; row <= layout.StoreLastRow; row++)
        {
            for (var column = grid.FirstColumn; column <= layout.TotalColumn; column++)
            {
                var formula = LoaderFormulaRepair.Expand(
                    grid.Formula(row, column), layout.FirstCodeColumn, layout.LastCodeColumn);

                if (formula is null)
                {
                    continue;
                }

                ExcelSheetOperations.SetFormula(sheet, row, column, formula);
                repaired++;
            }
        }

        _logger.Information(
            "На листе «" + PriceSchema.LoaderSheet + "» починено сумм по колонкам: " + repaired + ".");
    }

    /// <summary>
    /// Размножает первую колонку кодов по остальным. Все формулы колонки ссылаются на
    /// собственную ячейку кода (ВПР(G12; …)), поэтому копия сама начинает считать свой код.
    /// Коды пишутся после копирования - копия затирает строку заголовков.
    /// </summary>
    private void CopyFirstColumn(object application, object sheet, LoaderLayout layout)
    {
        if (layout.CodeColumnCount <= 1)
        {
            return;
        }

        for (var index = 1; index < layout.CodeColumnCount; index++)
        {
            ExcelSheetOperations.CopyRange(
                application,
                sheet,
                1,
                layout.FirstCodeColumn,
                layout.StoreLastRow,
                layout.FirstCodeColumn,
                1,
                layout.CodeColumnAt(index));
        }

        _logger.Information(
            "Первая колонка «" + PriceSchema.LoaderSheet + "» размножена на " +
            (layout.CodeColumnCount - 1) + " повтор(ов).");
    }

    /// <summary>
    /// Вписывает коды листа «Цены» в строку заголовков - это и есть «скопировать колонку
    /// «Код» и вставить транспонированием».
    /// </summary>
    private void WriteCodes(
        object sheet,
        LoaderLayout layout,
        IReadOnlyList<PriceRowValues> rows,
        List<ProcessingWarning> warnings)
    {
        var missing = 0;
        var codes = new object?[1, layout.CodeColumnCount];

        for (var index = 0; index < layout.CodeColumnCount; index++)
        {
            var code = index < rows.Count ? TextUtils.Normalize(rows[index].Reference?.Code) : string.Empty;
            if (code.Length == 0)
            {
                missing++;
                codes[0, index] = null;
                continue;
            }

            codes[0, index] = AsCode(code);
        }

        ExcelSheetOperations.SetBlockValues(sheet, layout.HeaderRow, layout.FirstCodeColumn, codes);

        if (missing > 0)
        {
            warnings.Add(new ProcessingWarning(
                "Строк без кода в поставке: " + missing + ". На листе «" + PriceSchema.LoaderSheet +
                "» их колонки остались пустыми - сначала нужно найти эти штрихкоды в «" +
                PriceSchema.SummaryPriceSheet + "».",
                PriceSchema.LoaderSheet + ", строка " + layout.HeaderRow,
                ExcelSheetOperations.GetSheetName(sheet),
                new CellRef(layout.HeaderRow, layout.FirstCodeColumn).ToString()));
        }

        _logger.Information(
            "Лист «" + PriceSchema.LoaderSheet + "»: колонок кодов " + layout.CodeColumnCount +
            ", из них без кода " + missing + ".");
    }

    /// <summary>
    /// Код пишется числом, если он число: на листе «Цены» он тоже число, а ВПР по
    /// текстовому коду в числовой колонке уже не найдёт ничего.
    /// </summary>
    private static object AsCode(string code) =>
        double.TryParse(code, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number)
            ? number
            : code;
}
