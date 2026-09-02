using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>Границы прямоугольника листа в абсолютных координатах Excel.</summary>
internal readonly record struct SheetBounds(int FirstRow, int FirstColumn, int RowCount, int ColumnCount)
{
    public int LastRow => FirstRow + RowCount - 1;

    public int LastColumn => FirstColumn + ColumnCount - 1;
}

/// <summary>
/// Низкоуровневые операции над листом Excel. Каждая операция берёт только те COM-объекты,
/// которые ей нужны, и освобождает их сразу же: временных неосвобождённых обёрток не остаётся.
/// Параметры объявлены как object, чтобы позднее связывание не распространялось на вызывающий код.
/// </summary>
internal static class ExcelSheetOperations
{
    public static object? FindSheet(object workbookObject, string name, ComScope scope)
    {
        dynamic workbook = workbookObject;
        var wanted = TextUtils.NormalizeKey(name);
        dynamic sheets = scope.Track(workbook.Worksheets);
        int count = sheets.Count;

        for (var index = 1; index <= count; index++)
        {
            dynamic sheet = sheets[index];
            string sheetName = sheet.Name;
            if (string.Equals(TextUtils.NormalizeKey(sheetName), wanted, StringComparison.Ordinal))
            {
                scope.Track(sheet);
                return (object)sheet;
            }

            ComUtils.Release(sheet);
        }

        return null;
    }

    public static SheetGrid ReadGrid(object sheetObject, bool withFormulas)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic used = scope.Track(sheet.UsedRange);
        int firstRow = used.Row;
        int firstColumn = used.Column;

        dynamic usedRows = scope.Track(used.Rows);
        int rowCount = usedRows.Count;
        dynamic usedColumns = scope.Track(used.Columns);
        int columnCount = usedColumns.Count;

        object? rawValues = used.Value2;
        object? rawFormulas = withFormulas ? used.Formula : null;

        var values = ToObjectArray(rawValues, rowCount, columnCount);
        var formulas = withFormulas ? ToStringArray(rawFormulas, rowCount, columnCount) : null;

        return new SheetGrid(firstRow, firstColumn, values, formulas);
    }

    /// <summary>Границы использованного диапазона листа. Данные при этом не передаются.</summary>
    public static SheetBounds GetUsedBounds(object sheetObject)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic used = scope.Track(sheet.UsedRange);
        int firstRow = used.Row;
        int firstColumn = used.Column;
        dynamic usedRows = scope.Track(used.Rows);
        int rowCount = usedRows.Count;
        dynamic usedColumns = scope.Track(used.Columns);
        int columnCount = usedColumns.Count;
        return new SheetBounds(firstRow, firstColumn, rowCount, columnCount);
    }

    /// <summary>
    /// Последняя заполненная строка колонки. Использованный диапазон Excel часто оказывается
    /// намного больше фактических данных, и без этой проверки пришлось бы читать пустые строки.
    /// </summary>
    public static int GetLastFilledRow(object sheetObject, int column, int fallbackLastRow)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic sheetRows = scope.Track(sheet.Rows);
        int sheetRowCount = sheetRows.Count;
        dynamic cells = scope.Track(sheet.Cells);
        dynamic bottom = scope.Track(cells[sheetRowCount, column]);
        dynamic last = scope.Track(bottom.End(ExcelConstants.XlUp));
        int lastRow = last.Row;
        return Math.Min(Math.Max(lastRow, 1), fallbackLastRow);
    }

    /// <summary>Читает прямоугольный фрагмент листа. Позволяет обрабатывать лист по частям.</summary>
    public static SheetGrid ReadBlock(
        object sheetObject,
        int firstRow,
        int lastRow,
        int firstColumn,
        int lastColumn,
        bool withFormulas)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        var reference =
            ExcelColumn.ToLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture) + ":" +
            ExcelColumn.ToLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture);
        dynamic range = scope.Track(sheet.Range[reference]);

        object? rawValues = range.Value2;
        object? rawFormulas = withFormulas ? range.Formula : null;

        var rowCount = lastRow - firstRow + 1;
        var columnCount = lastColumn - firstColumn + 1;

        return new SheetGrid(
            firstRow,
            firstColumn,
            ToObjectArray(rawValues, rowCount, columnCount),
            withFormulas ? ToStringArray(rawFormulas, rowCount, columnCount) : null);
    }

    public static string?[] ReadRowFormulas(object sheetObject, int row, int firstColumn, int lastColumn)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        var reference = BuildRangeReference(row, firstColumn, lastColumn);
        dynamic range = scope.Track(sheet.Range[reference]);
        object? raw = range.Formula;

        var count = lastColumn - firstColumn + 1;
        var result = new string?[count];
        if (raw is object[,] array)
        {
            var rowBound = array.GetLowerBound(0);
            var columnBound = array.GetLowerBound(1);
            for (var i = 0; i < count; i++)
            {
                result[i] = array[rowBound, columnBound + i] as string;
            }
        }
        else if (count == 1)
        {
            result[0] = raw as string;
        }

        return result;
    }

    public static void SetValue(object sheetObject, int row, int column, object? value)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        cell.Value2 = value;
    }

    public static void ClearValue(object sheetObject, int row, int column)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        cell.ClearContents();
    }

    /// <summary>
    /// Розовая заливка ячейки - стандартный цвет Excel «Плохо» (RGB 255 199 206).
    /// В COM цвет задаётся в порядке BGR, отсюда 0xCEC7FF.
    /// </summary>
    private const int WarningFill = 0xCEC7FF;

    /// <summary>
    /// Помечает или снимает пометку с ячейки. Снимается заливка только своя:
    /// если аналитик покрасила ячейку сама, её цвет сохраняется.
    /// </summary>
    public static void SetWarningFill(object sheetObject, int row, int column, bool highlight)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        dynamic interior = scope.Track(cell.Interior);

        if (highlight)
        {
            interior.Color = WarningFill;
            return;
        }

        object? current = interior.Color;
        if (current is not null && Convert.ToInt32(current, CultureInfo.InvariantCulture) == WarningFill)
        {
            interior.ColorIndex = ExcelConstants.XlColorIndexNone;
        }
    }

    public static void SetFormula(object sheetObject, int row, int column, string formula)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        cell.Formula = formula;
    }

    public static string? GetFormula(object sheetObject, int row, int column)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        object? formula = cell.Formula;
        return formula as string;
    }

    /// <summary>
    /// Снимает наложенные условия автофильтра, если они есть.
    ///
    /// Это обязательный шаг перед удалением и вставкой строк: при активном отборе
    /// Excel применяет Range.Delete не ко всему диапазону, а только к его видимой части,
    /// и часть строк остаётся на листе. Сам автофильтр (кнопки отбора) сохраняется.
    /// </summary>
    public static bool ShowAllRows(object sheetObject)
    {
        dynamic sheet = sheetObject;
        bool filterMode = sheet.FilterMode;
        if (!filterMode)
        {
            return false;
        }

        sheet.ShowAllData();
        return true;
    }

    /// <summary>
    /// Разделитель областей в адресе диапазона. Excel берёт его из настроек локали,
    /// поэтому на русской системе это «;», а не «,», и жёстко задавать его нельзя.
    /// </summary>
    public static string GetListSeparator(object applicationObject)
    {
        dynamic application = applicationObject;
        try
        {
            object? value = application.International(ExcelConstants.XlListSeparator);
            var separator = value as string;
            return string.IsNullOrEmpty(separator) ? "," : separator;
        }
        catch (RuntimeBinderException)
        {
            return ",";
        }
        catch (COMException)
        {
            return ",";
        }
    }

    /// <summary>
    /// Удаляет строки снизу вверх. Смежные строки объединяются в блоки, а блоки - в пачки:
    /// за один вызов Delete удаляется столько областей, сколько помещается в адрес диапазона
    /// (Excel ограничивает его 255 символами). На больших выгрузках это примерно вдвое быстрее,
    /// чем удаление каждого блока по отдельности.
    /// </summary>
    /// <summary>
    /// Удаляет строки. Если задан диапазон колонок, удаляется только прямоугольник внутри
    /// него со сдвигом вверх: на листе «План» иначе уехала бы боковая сводка справа
    /// от таблицы, а строка сводки, попавшая под удаление, пропала бы совсем.
    /// </summary>
    public static int DeleteRows(
        object sheetObject,
        IEnumerable<int> rows,
        string listSeparator,
        ColumnRange? columns = null)
    {
        dynamic sheet = sheetObject;
        var ordered = rows.Distinct().OrderBy(r => r).ToList();
        if (ordered.Count == 0)
        {
            return 0;
        }

        var blocks = BuildContiguousBlocks(ordered);
        const int addressBudget = 240;

        var index = blocks.Count - 1;
        while (index >= 0)
        {
            var parts = new List<string>();
            var length = 0;

            while (index >= 0)
            {
                var part = columns is null
                    ? RowReference(blocks[index].Start, blocks[index].End)
                    : BlockReference(blocks[index].Start, blocks[index].End, columns.Value);

                if (parts.Count > 0 && length + listSeparator.Length + part.Length > addressBudget)
                {
                    break;
                }

                parts.Add(part);
                length += part.Length + listSeparator.Length;
                index--;
            }

            using var scope = new ComScope();
            dynamic range = scope.Track(sheet.Range[string.Join(listSeparator, parts)]);

            if (columns is null)
            {
                range.Delete();
            }
            else
            {
                range.Delete(ExcelConstants.XlUp);
            }
        }

        return ordered.Count;
    }

    private static List<(int Start, int End)> BuildContiguousBlocks(IReadOnlyList<int> orderedRows)
    {
        var blocks = new List<(int Start, int End)>();
        var start = orderedRows[0];
        var previous = orderedRows[0];

        for (var i = 1; i < orderedRows.Count; i++)
        {
            if (orderedRows[i] == previous + 1)
            {
                previous = orderedRows[i];
                continue;
            }

            blocks.Add((start, previous));
            start = orderedRows[i];
            previous = orderedRows[i];
        }

        blocks.Add((start, previous));
        return blocks;
    }

    /// <summary>Копирует строку-шаблон и вставляет копию перед указанной строкой.</summary>
    public static void InsertCopiedRow(
        object applicationObject,
        object sheetObject,
        int templateRow,
        int insertBeforeRow,
        ColumnRange columns)
    {
        dynamic application = applicationObject;
        dynamic sheet = sheetObject;

        using (var scope = new ComScope())
        {
            dynamic source = scope.Track(sheet.Range[BlockReference(templateRow, templateRow, columns)]);
            source.Copy();
            dynamic destination = scope.Track(sheet.Range[BlockReference(insertBeforeRow, insertBeforeRow, columns)]);
            destination.Insert(ExcelConstants.XlDown);
        }

        application.CutCopyMode = false;
    }

    public static double GetRowHeight(object sheetObject, int row)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic rows = scope.Track(sheet.Rows);
        dynamic target = scope.Track(rows[row]);
        return Convert.ToDouble(target.RowHeight, CultureInfo.InvariantCulture);
    }

    public static void SetRowHeight(object sheetObject, int row, double height)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic rows = scope.Track(sheet.Rows);
        dynamic target = scope.Track(rows[row]);
        target.RowHeight = height;
    }

    /// <summary>Переносит строку целиком: вырезает и вставляет перед целевой строкой.</summary>
    public static void MoveRow(
        object applicationObject,
        object sheetObject,
        int fromRow,
        int toRow,
        ColumnRange columns)
    {
        dynamic application = applicationObject;
        dynamic sheet = sheetObject;

        using (var scope = new ComScope())
        {
            dynamic source = scope.Track(sheet.Range[BlockReference(fromRow, fromRow, columns)]);
            source.Cut();
            dynamic destination = scope.Track(sheet.Range[BlockReference(toRow, toRow, columns)]);
            destination.Insert(ExcelConstants.XlDown);
        }

        application.CutCopyMode = false;
    }

    private static string RowReference(int start, int end) =>
        start.ToString(CultureInfo.InvariantCulture) + ":" + end.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// Прямоугольник из строк, ограниченный колонками таблицы. Вставка, удаление и перенос
    /// строк выполняются именно так: если двигать строку целиком, вместе с ней уезжает всё,
    /// что стоит справа от таблицы - например, боковая сводка «норма в день / в план».
    /// </summary>
    private static string BlockReference(int startRow, int endRow, ColumnRange columns) =>
        ExcelColumn.ToLetters(columns.First) + startRow.ToString(CultureInfo.InvariantCulture) + ":" +
        ExcelColumn.ToLetters(columns.Last) + endRow.ToString(CultureInfo.InvariantCulture);

    private static string BuildRangeReference(int row, int firstColumn, int lastColumn) =>
        ExcelColumn.ToLetters(firstColumn) + row.ToString(CultureInfo.InvariantCulture) + ":" +
        ExcelColumn.ToLetters(lastColumn) + row.ToString(CultureInfo.InvariantCulture);

    private static object?[,] ToObjectArray(object? raw, int rowCount, int columnCount)
    {
        var result = new object?[rowCount, columnCount];
        if (raw is object[,] array)
        {
            var rowBound = array.GetLowerBound(0);
            var columnBound = array.GetLowerBound(1);
            for (var r = 0; r < rowCount; r++)
            {
                for (var c = 0; c < columnCount; c++)
                {
                    result[r, c] = array[rowBound + r, columnBound + c];
                }
            }
        }
        else if (raw is not null && rowCount == 1 && columnCount == 1)
        {
            result[0, 0] = raw;
        }

        return result;
    }

    private static string?[,] ToStringArray(object? raw, int rowCount, int columnCount)
    {
        var result = new string?[rowCount, columnCount];
        if (raw is object[,] array)
        {
            var rowBound = array.GetLowerBound(0);
            var columnBound = array.GetLowerBound(1);
            for (var r = 0; r < rowCount; r++)
            {
                for (var c = 0; c < columnCount; c++)
                {
                    result[r, c] = array[rowBound + r, columnBound + c] as string;
                }
            }
        }
        else if (raw is not null && rowCount == 1 && columnCount == 1)
        {
            result[0, 0] = raw as string;
        }

        return result;
    }
}
