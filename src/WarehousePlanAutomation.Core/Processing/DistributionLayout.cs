using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Разметка листа «Распред».
///
/// Строки заголовков в привычном смысле у листа нет, поэтому разметка читается двумя
/// способами: подписи («распред», «в загрузку», «Ост», «Ед») ищутся по тексту, а блок
/// сезонности и ячейка сектора - в формулах таблицы РТТ, где они и используются.
/// Так ни один адрес не приходится задавать числом.
/// </summary>
public sealed class DistributionLayout
{
    private DistributionLayout(
        int headerRow,
        int labelColumn,
        int firstBlockColumn,
        int blockCount,
        int storeFirstRow,
        int storeLastRow,
        int codeRow,
        int unitsRow,
        int remainderRow,
        int hubMinimumRow,
        RangeRef seasonality,
        CellRef sectorCell)
    {
        HeaderRow = headerRow;
        LabelColumn = labelColumn;
        FirstBlockColumn = firstBlockColumn;
        BlockCount = blockCount;
        StoreFirstRow = storeFirstRow;
        StoreLastRow = storeLastRow;
        CodeRow = codeRow;
        UnitsRow = unitsRow;
        RemainderRow = remainderRow;
        HubMinimumRow = hubMinimumRow;
        Seasonality = seasonality;
        SectorCell = sectorCell;
    }

    /// <summary>Строка заголовков таблицы РТТ: «Код», «Концепт», …, «распред», «в загрузку».</summary>
    public int HeaderRow { get; }

    /// <summary>Колонка подписей слева от блоков: «Код», «Ед», «Ост», «мин запас на Хаб».</summary>
    public int LabelColumn { get; }

    public int FirstBlockColumn { get; }

    public int BlockCount { get; }

    public int StoreFirstRow { get; }

    public int StoreLastRow { get; }

    public int CodeRow { get; }

    public int UnitsRow { get; }

    public int RemainderRow { get; }

    public int HubMinimumRow { get; }

    /// <summary>Блок сезонности: подгруппа, месяцы, «Мин раздача». Взят из формулы «распред».</summary>
    public RangeRef Seasonality { get; }

    /// <summary>Ячейка с названием сектора. Взята из формулы таблицы РТТ.</summary>
    public CellRef SectorCell { get; }

    public int LastBlockColumn => FirstBlockColumn + (BlockCount * PriceSchema.Distribution.BlockWidth) - 1;

    /// <summary>
    /// Та же разметка, но с другим числом блоков.
    ///
    /// Нужна сразу после вставки колонок: на листе в этот момент стоят пустые колонки,
    /// подписей «в загрузку» в них ещё нет, и пересчитанное по листу число блоков было бы
    /// меньше настоящего. Сколько блоков должно получиться, знает только вызывающий.
    /// </summary>
    public DistributionLayout WithBlockCount(int blockCount) =>
        new(HeaderRow, LabelColumn, FirstBlockColumn, blockCount, StoreFirstRow, StoreLastRow,
            CodeRow, UnitsRow, RemainderRow, HubMinimumRow, Seasonality, SectorCell);

    /// <summary>Первая колонка блока с номером <paramref name="index"/>, считая с нуля.</summary>
    public int BlockColumn(int index) =>
        FirstBlockColumn + (index * PriceSchema.Distribution.BlockWidth);

    /// <summary>Строка с названиями месяцев - прямо над блоком сезонности.</summary>
    public int MonthRow => Seasonality.FirstRow - 1;

    public int SubgroupColumn => Seasonality.FirstColumn;

    /// <summary>Сколько месяцев помещается в блок: без колонки подгруппы и «Мин раздачи».</summary>
    public int MonthCount => Math.Max(Seasonality.ColumnCount - 2, 0);

    public int MonthColumn(int index) => Seasonality.FirstColumn + 1 + index;

    /// <summary>Сколько подгрупп помещается в блок сезонности.</summary>
    public int SubgroupCapacity => Seasonality.RowCount;

    public static DistributionLayout Read(SheetGrid grid)
    {
        var problems = new List<string>();

        var headerRow = FindHeaderRow(grid);
        if (headerRow is null)
        {
            problems.Add("на листе «" + PriceSchema.DistributionSheet +
                         "» не найдена строка заголовков таблицы (ожидается «в загрузку»)");
            throw new WorkbookValidationException(problems);
        }

        var firstBlockColumn = FindFirstBlockColumn(grid, headerRow.Value);
        var blockCount = CountBlocks(grid, headerRow.Value);

        if (firstBlockColumn is null || blockCount == 0)
        {
            problems.Add("на листе «" + PriceSchema.DistributionSheet +
                         "» не найден ни один блок АЦР (ожидаются колонки «распред» и «в загрузку»)");
            throw new WorkbookValidationException(problems);
        }

        var labelColumn = firstBlockColumn.Value - 1;

        int? RowOfLabel(string label, bool startsWith = false)
        {
            for (var row = grid.FirstRow; row < headerRow.Value; row++)
            {
                var text = TextUtils.NormalizeKey(grid.Text(row, labelColumn));
                if (startsWith
                        ? text.StartsWith(label, StringComparison.Ordinal)
                        : string.Equals(text, label, StringComparison.Ordinal))
                {
                    return row;
                }
            }

            return null;
        }

        var codeRow = RowOfLabel(PriceSchema.Distribution.LabelCode);
        var unitsRow = RowOfLabel(PriceSchema.Distribution.LabelUnits);
        var remainderRow = RowOfLabel(PriceSchema.Distribution.LabelRemainder);
        var hubMinimumRow = RowOfLabel(PriceSchema.Distribution.LabelHubMinimum, startsWith: true);

        void Require(int? row, string label)
        {
            if (row is null)
            {
                problems.Add("на листе «" + PriceSchema.DistributionSheet + "» не найдена подпись «" +
                             label + "» в колонке " + ExcelColumn.ToLetters(labelColumn));
            }
        }

        Require(codeRow, PriceSchema.Distribution.LabelCode);
        Require(unitsRow, PriceSchema.Distribution.LabelUnits);
        Require(remainderRow, PriceSchema.Distribution.LabelRemainder);
        Require(hubMinimumRow, PriceSchema.Distribution.LabelHubMinimum);

        var storeFirstRow = headerRow.Value + 1;
        var storeLastRow = FindStoreLastRow(grid, storeFirstRow);
        if (storeLastRow < storeFirstRow)
        {
            problems.Add("на листе «" + PriceSchema.DistributionSheet + "» нет ни одной строки РТТ");
        }

        var seasonality = ExcelReference.FindLocalRange(
            grid.Formula(storeFirstRow, firstBlockColumn.Value));
        if (seasonality is null)
        {
            problems.Add("на листе «" + PriceSchema.DistributionSheet +
                         "» не удалось определить блок сезонности: в формуле колонки «" +
                         PriceSchema.Distribution.BlockDistribute + "» нет ссылки на него");
        }

        var sectorCell = FindSectorCell(grid, storeFirstRow, labelColumn);
        if (sectorCell is null)
        {
            problems.Add("на листе «" + PriceSchema.DistributionSheet +
                         "» не удалось определить ячейку сектора: в формулах таблицы РТТ нет ключа вида " +
                         "ТЕКСТ(код)&ячейка");
        }

        if (problems.Count > 0)
        {
            throw new WorkbookValidationException(problems);
        }

        return new DistributionLayout(
            headerRow.Value,
            labelColumn,
            firstBlockColumn.Value,
            blockCount,
            storeFirstRow,
            storeLastRow,
            codeRow!.Value,
            unitsRow!.Value,
            remainderRow!.Value,
            hubMinimumRow!.Value,
            seasonality!.Value,
            sectorCell!.Value);
    }

    private static int? FindHeaderRow(SheetGrid grid)
    {
        for (var row = grid.FirstRow; row <= grid.LastRow; row++)
        {
            for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
            {
                if (TextUtils.EqualsKey(grid.Text(row, column), PriceSchema.Distribution.BlockShip))
                {
                    return row;
                }
            }
        }

        return null;
    }

    private static int? FindFirstBlockColumn(SheetGrid grid, int headerRow)
    {
        for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
        {
            if (TextUtils.EqualsKey(grid.Text(headerRow, column), PriceSchema.Distribution.BlockDistribute))
            {
                return column;
            }
        }

        return null;
    }

    private static int CountBlocks(SheetGrid grid, int headerRow)
    {
        var count = 0;
        for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
        {
            if (TextUtils.EqualsKey(grid.Text(headerRow, column), PriceSchema.Distribution.BlockShip))
            {
                count++;
            }
        }

        return count;
    }

    /// <summary>Строки РТТ идут подряд и заканчиваются там, где кончаются коды в первой колонке.</summary>
    private static int FindStoreLastRow(SheetGrid grid, int firstRow)
    {
        var last = firstRow - 1;
        for (var row = firstRow; row <= grid.LastRow; row++)
        {
            if (TextUtils.Normalize(grid.Text(row, grid.FirstColumn)).Length == 0)
            {
                break;
            }

            last = row;
        }

        return last;
    }

    /// <summary>
    /// Ключ поиска в таблице РТТ склеен из кода магазина и сектора: TEXT($A28;"000")&amp;$H$26.
    /// Ячейка сектора берётся из этой формулы, а не по расположению.
    /// </summary>
    private static CellRef? FindSectorCell(SheetGrid grid, int storeRow, int labelColumn)
    {
        for (var column = grid.FirstColumn; column <= labelColumn; column++)
        {
            var cell = ExcelReference.FindConcatenatedCell(grid.Formula(storeRow, column));
            if (cell is not null)
            {
                return cell;
            }
        }

        return null;
    }
}
