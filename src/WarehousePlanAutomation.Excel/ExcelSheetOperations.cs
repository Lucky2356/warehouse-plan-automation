using System.Globalization;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

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

    /// <summary>Удаляет строки снизу вверх непрерывными блоками.</summary>
    public static int DeleteRows(object sheetObject, IEnumerable<int> rows)
    {
        dynamic sheet = sheetObject;
        var ordered = rows.Distinct().OrderBy(r => r).ToList();
        if (ordered.Count == 0)
        {
            return 0;
        }

        var blocks = new List<(int Start, int End)>();
        var start = ordered[0];
        var previous = ordered[0];

        for (var i = 1; i < ordered.Count; i++)
        {
            if (ordered[i] == previous + 1)
            {
                previous = ordered[i];
                continue;
            }

            blocks.Add((start, previous));
            start = ordered[i];
            previous = ordered[i];
        }

        blocks.Add((start, previous));

        for (var i = blocks.Count - 1; i >= 0; i--)
        {
            using var scope = new ComScope();
            dynamic range = scope.Track(sheet.Range[RowReference(blocks[i].Start, blocks[i].End)]);
            range.Delete();
        }

        return ordered.Count;
    }

    /// <summary>Копирует строку-шаблон и вставляет копию перед указанной строкой.</summary>
    public static void InsertCopiedRow(
        object applicationObject,
        object sheetObject,
        int templateRow,
        int insertBeforeRow)
    {
        dynamic application = applicationObject;
        dynamic sheet = sheetObject;

        using (var scope = new ComScope())
        {
            dynamic source = scope.Track(sheet.Range[RowReference(templateRow, templateRow)]);
            source.Copy();
            dynamic destination = scope.Track(sheet.Range[RowReference(insertBeforeRow, insertBeforeRow)]);
            destination.Insert(ExcelConstants.XlDown);
        }

        application.CutCopyMode = false;
    }

    /// <summary>Переносит строку целиком: вырезает и вставляет перед целевой строкой.</summary>
    public static void MoveRow(object applicationObject, object sheetObject, int fromRow, int toRow)
    {
        dynamic application = applicationObject;
        dynamic sheet = sheetObject;

        using (var scope = new ComScope())
        {
            dynamic source = scope.Track(sheet.Range[RowReference(fromRow, fromRow)]);
            source.Cut();
            dynamic destination = scope.Track(sheet.Range[RowReference(toRow, toRow)]);
            destination.Insert(ExcelConstants.XlDown);
        }

        application.CutCopyMode = false;
    }

    private static string RowReference(int start, int end) =>
        start.ToString(CultureInfo.InvariantCulture) + ":" + end.ToString(CultureInfo.InvariantCulture);

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
