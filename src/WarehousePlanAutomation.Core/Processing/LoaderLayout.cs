using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Разметка листа «Загрузочник».
///
/// Лист устроен как таблица РТТ: слева неизменные колонки («Концепт», «Доля», «Хаб»,
/// «Код»), дальше по колонке на каждый код поставки, и замыкает всё «Итого». Над
/// таблицей стоит блок формул по тем же колонкам, а слева от него - ячейка проверки,
/// та самая, где должно получиться «ИСТИНА».
///
/// Ни одна буква колонки и ни один номер строки не зашиты: строка заголовков находится
/// по паре подписей «Код» … «Итого», а ячейка проверки - по виду формулы (сравнение),
/// потому что подписи у неё нет.
/// </summary>
public sealed class LoaderLayout
{
    private LoaderLayout(
        int headerRow,
        int codeColumn,
        int codeColumnCount,
        int totalColumn,
        int storeFirstRow,
        int storeLastRow,
        CellRef checkCell)
    {
        HeaderRow = headerRow;
        CodeColumn = codeColumn;
        CodeColumnCount = codeColumnCount;
        TotalColumn = totalColumn;
        StoreFirstRow = storeFirstRow;
        StoreLastRow = storeLastRow;
        CheckCell = checkCell;
    }

    /// <summary>Строка с подписями «Код», …, кодами поставки и «Итого».</summary>
    public int HeaderRow { get; }

    /// <summary>Колонка «Код»: в ней коды РТТ, а не коды товара.</summary>
    public int CodeColumn { get; }

    /// <summary>Сколько колонок сейчас отведено под коды поставки.</summary>
    public int CodeColumnCount { get; }

    /// <summary>Колонка «Итого» - сразу за колонками кодов.</summary>
    public int TotalColumn { get; }

    public int StoreFirstRow { get; }

    public int StoreLastRow { get; }

    /// <summary>Ячейка, в которой сходятся «Распред» и «Загрузочник». Должна дать «ИСТИНА».</summary>
    public CellRef CheckCell { get; }

    /// <summary>Первая колонка кодов поставки.</summary>
    public int FirstCodeColumn => CodeColumn + 1;

    public int LastCodeColumn => FirstCodeColumn + CodeColumnCount - 1;

    /// <summary>Колонка кода с номером <paramref name="index"/>, считая с нуля.</summary>
    public int CodeColumnAt(int index) => FirstCodeColumn + index;

    /// <summary>
    /// Та же разметка, но с другим числом колонок кодов. Нужна сразу после вставки:
    /// новые колонки пока пустые, «Итого» уже сдвинулось, и пересчитывать разметку
    /// по листу рано.
    /// </summary>
    public LoaderLayout WithCodeColumnCount(int count) =>
        new(HeaderRow, CodeColumn, count, FirstCodeColumn + count, StoreFirstRow, StoreLastRow, CheckCell);

    public static LoaderLayout Read(SheetGrid grid)
    {
        var problems = new List<string>();

        var header = FindHeader(grid);
        if (header is null)
        {
            problems.Add("на листе «" + PriceSchema.LoaderSheet +
                         "» не найдена строка заголовков: в одной строке должны стоять «" +
                         PriceSchema.Loader.LabelCode + "» и правее «" +
                         PriceSchema.Loader.LabelTotal + "»");
            throw new WorkbookValidationException(problems);
        }

        var (headerRow, codeColumn, totalColumn) = header.Value;
        var codeColumnCount = totalColumn - codeColumn - 1;

        if (codeColumnCount < 1)
        {
            problems.Add("на листе «" + PriceSchema.LoaderSheet + "» между «" +
                         PriceSchema.Loader.LabelCode + "» и «" + PriceSchema.Loader.LabelTotal +
                         "» нет ни одной колонки кода");
        }

        var storeFirstRow = headerRow + 1;
        var storeLastRow = FindStoreLastRow(grid, storeFirstRow, codeColumn);
        if (storeLastRow < storeFirstRow)
        {
            problems.Add("на листе «" + PriceSchema.LoaderSheet + "» нет ни одной строки РТТ");
        }

        var checkCell = FindCheckCell(grid, headerRow, codeColumn);
        if (checkCell is null)
        {
            problems.Add("на листе «" + PriceSchema.LoaderSheet +
                         "» не найдена ячейка проверки: над таблицей нет формулы-сравнения, " +
                         "которая должна давать «ИСТИНА»");
        }

        if (problems.Count > 0)
        {
            throw new WorkbookValidationException(problems);
        }

        return new LoaderLayout(
            headerRow, codeColumn, codeColumnCount, totalColumn,
            storeFirstRow, storeLastRow, checkCell!.Value);
    }

    /// <summary>Строка заголовков: «Код» и правее него «Итого».</summary>
    private static (int Row, int CodeColumn, int TotalColumn)? FindHeader(SheetGrid grid)
    {
        for (var row = grid.FirstRow; row <= grid.LastRow; row++)
        {
            for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
            {
                if (!TextUtils.EqualsKey(grid.Text(row, column), PriceSchema.Loader.LabelCode))
                {
                    continue;
                }

                for (var right = column + 1; right <= grid.LastColumn; right++)
                {
                    if (TextUtils.EqualsKey(grid.Text(row, right), PriceSchema.Loader.LabelTotal))
                    {
                        return (row, column, right);
                    }
                }
            }
        }

        return null;
    }

    /// <summary>Строки РТТ идут подряд и заканчиваются там, где кончаются коды магазинов.</summary>
    private static int FindStoreLastRow(SheetGrid grid, int firstRow, int codeColumn)
    {
        var last = firstRow - 1;
        for (var row = firstRow; row <= grid.LastRow; row++)
        {
            if (TextUtils.Normalize(grid.Text(row, codeColumn)).Length == 0)
            {
                break;
            }

            last = row;
        }

        return last;
    }

    /// <summary>
    /// Ячейка проверки подписи не имеет, зато отличается формулой: это единственное
    /// сравнение над таблицей. Искать её по виду формулы надёжнее, чем по расположению.
    /// </summary>
    private static CellRef? FindCheckCell(SheetGrid grid, int headerRow, int codeColumn)
    {
        for (var row = grid.FirstRow; row < headerRow; row++)
        {
            for (var column = grid.FirstColumn; column < codeColumn; column++)
            {
                if (IsComparison(grid.Formula(row, column)))
                {
                    return new CellRef(row, column);
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Формула-сравнение: знак сравнения вне скобок и вне текста. Внутри скобок такие
    /// знаки встречаются у обычных функций (ЕСЛИ, СУММЕСЛИ), поэтому они не в счёт.
    /// </summary>
    public static bool IsComparison(string? formula)
    {
        if (string.IsNullOrEmpty(formula) || formula[0] != '=')
        {
            return false;
        }

        var depth = 0;
        for (var i = 1; i < formula.Length; i++)
        {
            var symbol = formula[i];

            if (symbol == '"')
            {
                i = formula.IndexOf('"', i + 1);
                if (i < 0)
                {
                    return false;
                }

                continue;
            }

            if (symbol == '(')
            {
                depth++;
            }
            else if (symbol == ')')
            {
                depth--;
            }
            else if (depth == 0 && (symbol == '=' || symbol == '<' || symbol == '>'))
            {
                return true;
            }
        }

        return false;
    }
}
