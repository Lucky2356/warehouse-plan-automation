using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Результат поиска листа. <see cref="Sheet"/> заполнен, когда подошёл ровно один лист;
/// иначе <see cref="Candidates"/> перечисляет неоднозначные названия (пустой список -
/// не нашлось ни одного).
/// </summary>
internal sealed record SheetLookup(object? Sheet, IReadOnlyList<string> Candidates);

/// <summary>Пометка строки «Плана» цветом.</summary>
internal enum RowMark
{
    /// <summary>Пометки нет; своя заливка снимается.</summary>
    None,

    /// <summary>Строка добавлена сегодня - зелёная.</summary>
    Added,

    /// <summary>Строку нужно разобрать вручную - розовая.</summary>
    Attention,
}

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
    /// <summary>
    /// Ищет лист по началу названия: аналитик дописывает к нему дату или пометку
    /// («Заказы на отгрузку 04_09», «Журнал заказов на отгрузку тест»), и это правильное
    /// название. Точное совпадение сильнее: если лист назван ровно так, как в схеме,
    /// берётся он, даже когда есть листы с более длинными названиями.
    ///
    /// Несколько подходящих листов - не повод угадывать: такой случай возвращается
    /// вызывающему как список названий, чтобы он объяснил его человеку.
    /// </summary>
    public static SheetLookup FindSheet(object workbookObject, string name, ComScope scope)
    {
        dynamic workbook = workbookObject;
        var wanted = TextUtils.NormalizeKey(name);
        dynamic sheets = scope.Track(workbook.Worksheets);
        int count = sheets.Count;

        var names = new List<string>(count);
        for (var index = 1; index <= count; index++)
        {
            dynamic sheet = sheets[index];
            string sheetName = sheet.Name;
            names.Add(sheetName);
            ComUtils.Release(sheet);
        }

        var exact = new List<int>();
        var byPrefix = new List<int>();
        for (var index = 0; index < names.Count; index++)
        {
            var key = TextUtils.NormalizeKey(names[index]);
            if (string.Equals(key, wanted, StringComparison.Ordinal))
            {
                exact.Add(index);
            }
            else if (key.StartsWith(wanted, StringComparison.Ordinal))
            {
                byPrefix.Add(index);
            }
        }

        var matches = exact.Count > 0 ? exact : byPrefix;
        if (matches.Count != 1)
        {
            return new SheetLookup(null, matches.Select(index => names[index]).ToList());
        }

        dynamic found = scope.Track(sheets[matches[0] + 1]);
        return new SheetLookup((object)found, Array.Empty<string>());
    }

    /// <summary>
    /// Все листы, название которых начинается с заданного, в порядке книги.
    /// Нужно там, где листов заведомо может быть несколько: инвойс приходит одним листом
    /// или двумя, и оба одинаково правильные.
    /// </summary>
    public static IReadOnlyList<object> FindSheets(object workbookObject, string name, ComScope scope)
    {
        dynamic workbook = workbookObject;
        var wanted = TextUtils.NormalizeKey(name);
        dynamic sheets = scope.Track(workbook.Worksheets);
        int count = sheets.Count;

        var found = new List<object>();
        for (var index = 1; index <= count; index++)
        {
            dynamic sheet = sheets[index];
            string sheetName = sheet.Name;
            if (TextUtils.NormalizeKey(sheetName).StartsWith(wanted, StringComparison.Ordinal))
            {
                scope.Track(sheet);
                found.Add((object)sheet);
                continue;
            }

            ComUtils.Release(sheet);
        }

        return found;
    }

    /// <summary>
    /// Удаляет лист с таким названием, если он есть. Предупреждение Excel «лист будет
    /// удалён навсегда» при этом отключается: иначе обработка встала бы на диалоге.
    /// </summary>
    public static bool RemoveSheet(object workbookObject, string name, ComScope scope)
    {
        dynamic workbook = workbookObject;
        var wanted = TextUtils.NormalizeKey(name);
        dynamic sheets = scope.Track(workbook.Worksheets);
        int count = sheets.Count;

        for (var index = count; index >= 1; index--)
        {
            dynamic sheet = sheets[index];
            string sheetName = sheet.Name;
            if (TextUtils.EqualsKey(sheetName, wanted))
            {
                sheet.Delete();
                ComUtils.Release(sheet);
                return true;
            }

            ComUtils.Release(sheet);
        }

        return false;
    }

    /// <summary>Добавляет пустой лист сразу после указанного и даёт ему имя.</summary>
    public static object AddSheet(object workbookObject, object afterSheetObject, string name, ComScope scope)
    {
        dynamic workbook = workbookObject;
        dynamic sheets = scope.Track(workbook.Worksheets);
        dynamic created = scope.Track(sheets.Add(Type.Missing, afterSheetObject));
        created.Name = name;
        return (object)created;
    }

    /// <summary>Добавляет пустой лист перед указанным и даёт ему имя.</summary>
    public static object AddSheetBefore(object workbookObject, object beforeSheetObject, string name, ComScope scope)
    {
        dynamic workbook = workbookObject;
        dynamic sheets = scope.Track(workbook.Worksheets);
        dynamic created = scope.Track(sheets.Add(beforeSheetObject));
        created.Name = name;
        return (object)created;
    }

    public static string GetSheetName(object sheetObject)
    {
        dynamic sheet = sheetObject;
        string name = sheet.Name;
        return name;
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
    /// Розовая заливка - стандартный цвет Excel «Плохо» (RGB 255 199 206).
    /// В COM цвет задаётся в порядке BGR, отсюда 0xCEC7FF.
    /// </summary>
    private const int AttentionFill = 0xCEC7FF;

    /// <summary>
    /// Зелёная заливка - парный к нему цвет Excel «Хорошо» (RGB 198 239 206),
    /// в порядке BGR 0xCEEFC6.
    /// </summary>
    private const int AddedFill = 0xCEEFC6;

    /// <summary>
    /// Помечает или снимает пометку с ячейки. Снимаются только свои заливки - две пометки
    /// и цвета повторов штрихкода: если аналитик покрасила ячейку сама, её цвет сохраняется.
    /// </summary>
    public static void SetRowMark(object sheetObject, int row, int column, RowMark mark)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        dynamic interior = scope.Track(cell.Interior);

        if (mark != RowMark.None)
        {
            interior.Color = mark == RowMark.Added ? AddedFill : AttentionFill;
            return;
        }

        object? current = interior.Color;
        if (current is null)
        {
            return;
        }

        var color = Convert.ToInt32(current, CultureInfo.InvariantCulture);
        if (color == AttentionFill || color == AddedFill ||
            WarehousePlanAutomation.Core.Processing.DuplicateBarcodeGroups.Palette.Contains(color))
        {
            interior.ColorIndex = ExcelConstants.XlColorIndexNone;
        }
    }

    /// <summary>
    /// Стоит ли в ячейке процентный формат. От этого зависит, что значит записанное
    /// в ней число: при формате «0%» 80 % лежит в ячейке как 0,8, а без него - как 80.
    /// </summary>
    public static bool IsPercentFormat(object sheetObject, int row, int column)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);

        object? format = cell.NumberFormat;
        return format is string text && text.Contains('%', StringComparison.Ordinal);
    }

    /// <summary>Задаёт числовой формат колонке: перенесённые значения должны читаться так же.</summary>
    public static void SetColumnFormat(object sheetObject, int column, string format)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic columns = scope.Track(sheet.Columns);
        dynamic target = scope.Track(columns[column]);
        target.NumberFormat = format;
    }

    /// <summary>Подгоняет ширину колонок под содержимое.</summary>
    public static void AutoFitColumns(object sheetObject, int firstColumn, int lastColumn)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic range = scope.Track(sheet.Range[
            BuildReference(1, firstColumn, 1, lastColumn)]);
        dynamic columns = scope.Track(range.EntireColumn);
        columns.AutoFit();
    }

    public static void SetFormula(object sheetObject, int row, int column, string formula)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        cell.Formula = formula;
    }

    /// <summary>
    /// Формула в относительной записи: перенесённая в другую строку, она считает уже свою
    /// строку - так же, как при протягивании.
    /// </summary>
    public static string? GetFormulaR1C1(object sheetObject, int row, int column)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        object? formula = cell.FormulaR1C1;
        return formula as string;
    }

    public static void SetFormulaR1C1(object sheetObject, int row, int column, string formula)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        cell.FormulaR1C1 = formula;
    }

    /// <summary>
    /// Красит часть текста ячейки. <paramref name="start"/> - позиция с нуля, как в строке C#;
    /// Excel считает символы с единицы. Цвет в порядке BGR.
    /// </summary>
    public static void SetTextColor(object sheetObject, int row, int column, int start, int length, int color)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        dynamic characters = scope.Track(cell.Characters[start + 1, length]);
        dynamic font = scope.Track(characters.Font);
        font.Color = color;
    }

    /// <summary>Красит заливку ячейки. Цвет в порядке BGR.</summary>
    public static void SetFill(object sheetObject, int row, int column, int color)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        dynamic interior = scope.Track(cell.Interior);
        interior.Color = color;
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

    /// <summary>
    /// Вставляет несколько копий строки-образца перед указанной строкой за один приём.
    /// Excel размножает одну скопированную строку по всему диапазону вставки, поэтому
    /// сто новых строк стоят одного обмена с COM, а не ста.
    /// </summary>
    public static void InsertCopiedRows(
        object applicationObject,
        object sheetObject,
        int templateRow,
        int insertBeforeRow,
        int count,
        ColumnRange columns)
    {
        if (count <= 0)
        {
            return;
        }

        dynamic application = applicationObject;
        dynamic sheet = sheetObject;

        using (var scope = new ComScope())
        {
            dynamic source = scope.Track(sheet.Range[BlockReference(templateRow, templateRow, columns)]);
            source.Copy();
            dynamic destination = scope.Track(
                sheet.Range[BlockReference(insertBeforeRow, insertBeforeRow + count - 1, columns)]);
            destination.Insert(ExcelConstants.XlDown);
        }

        application.CutCopyMode = false;
    }

    /// <summary>
    /// Записывает колонку целиком одним обменом с COM. По ячейке за раз лист на несколько
    /// сотен строк заполнялся бы заметно дольше, чем читается.
    /// </summary>
    public static void SetColumnValues(object sheetObject, int firstRow, int column, IReadOnlyList<object?> values)
    {
        if (values.Count == 0)
        {
            return;
        }

        dynamic sheet = sheetObject;
        using var scope = new ComScope();

        var letters = ExcelColumn.ToLetters(column);
        var reference =
            letters + firstRow.ToString(CultureInfo.InvariantCulture) + ":" +
            letters + (firstRow + values.Count - 1).ToString(CultureInfo.InvariantCulture);

        var block = new object?[values.Count, 1];
        for (var i = 0; i < values.Count; i++)
        {
            block[i, 0] = values[i];
        }

        dynamic range = scope.Track(sheet.Range[reference]);
        range.Value2 = block;
    }

    /// <summary>Записывает прямоугольник значений одним обменом с COM.</summary>
    public static void SetBlockValues(object sheetObject, int firstRow, int firstColumn, object?[,] values)
    {
        var rows = values.GetLength(0);
        var columns = values.GetLength(1);
        if (rows == 0 || columns == 0)
        {
            return;
        }

        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic range = scope.Track(sheet.Range[BuildReference(
            firstRow, firstColumn, firstRow + rows - 1, firstColumn + columns - 1)]);
        range.Value2 = values;
    }

    /// <summary>
    /// Записывает прямоугольник формул. Пустые значения очищают ячейку: так в одну
    /// операцию ставятся формулы там, где они нужны, и снимаются там, где нет.
    /// </summary>
    public static void SetBlockFormulas(object sheetObject, int firstRow, int firstColumn, object?[,] formulas)
    {
        var rows = formulas.GetLength(0);
        var columns = formulas.GetLength(1);
        if (rows == 0 || columns == 0)
        {
            return;
        }

        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic range = scope.Track(sheet.Range[BuildReference(
            firstRow, firstColumn, firstRow + rows - 1, firstColumn + columns - 1)]);
        range.Formula = formulas;
    }

    /// <summary>
    /// Переводит прямоугольник в текстовый формат. Нужно перед записью строк, которые
    /// Excel иначе разберёт как число: перечень магазинов «150,155,156,…» он принимает
    /// за число с разделителями разрядов.
    /// </summary>
    public static void SetRangeAsText(
        object sheetObject, int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic range = scope.Track(sheet.Range[BuildReference(firstRow, firstColumn, lastRow, lastColumn)]);
        range.NumberFormat = "@";
    }

    /// <summary>Читает значение одной ячейки.</summary>
    public static object? GetValue(object sheetObject, int row, int column)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic cells = scope.Track(sheet.Cells);
        dynamic cell = scope.Track(cells[row, column]);
        return cell.Value2;
    }

    /// <summary>Очищает значения в прямоугольнике, оставляя оформление.</summary>
    public static void ClearRange(
        object sheetObject, int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        if (lastRow < firstRow || lastColumn < firstColumn)
        {
            return;
        }

        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic range = scope.Track(sheet.Range[BuildReference(firstRow, firstColumn, lastRow, lastColumn)]);
        range.ClearContents();
    }

    /// <summary>
    /// Находит колонку с одним из заголовков <paramref name="keys"/>, а если её нет -
    /// дописывает новую сразу за последним заполненным заголовком. Оформление заголовка
    /// копируется из колонки <paramref name="formatSource"/>; если её нет (0), берётся
    /// последний заголовок строки - новая колонка выглядит так же, как соседние.
    /// </summary>
    /// <returns>Номер найденной или созданной колонки.</returns>
    public static int EnsureHeaderColumn(
        object sheetObject,
        int headerRow,
        int knownLastColumn,
        string title,
        IReadOnlyList<string> keys,
        int formatSource = 0)
    {
        var bounds = GetUsedBounds(sheetObject);
        var lastColumn = Math.Max(bounds.FirstColumn + bounds.ColumnCount - 1, knownLastColumn);
        var grid = ReadBlock(sheetObject, headerRow, headerRow, 1, lastColumn, withFormulas: false);

        var lastHeader = 0;
        for (var column = 1; column <= lastColumn; column++)
        {
            var key = TextUtils.NormalizeKey(grid.Text(headerRow, column));
            if (key.Length == 0)
            {
                continue;
            }

            lastHeader = column;
            if (keys.Contains(key, StringComparer.Ordinal))
            {
                return column;
            }
        }

        var created = lastHeader + 1;
        using (var scope = new ComScope())
        {
            dynamic sheet = sheetObject;
            dynamic from = scope.Track(sheet.Cells[headerRow, formatSource > 0 ? formatSource : Math.Max(lastHeader, 1)]);
            dynamic to = scope.Track(sheet.Cells[headerRow, created]);
            from.Copy(to);
        }

        SetValue(sheetObject, headerRow, created, title);
        return created;
    }

    /// <summary>
    /// Дописывает недостающие колонки в строку заголовков, сохраняя порядок списка:
    /// пропущенная колонка вставляется перед ближайшей следующей из списка, а если такой
    /// нет - приписывается за последним заголовком. Оформление берётся у соседнего заголовка.
    ///
    /// Поиск и вставка идут начиная с <paramref name="firstColumn"/>: на листах поставок
    /// первые колонки заняты формулами, и «Код» из них нельзя спутать с «Кодом» данных.
    /// </summary>
    /// <returns>Названия созданных колонок в порядке создания.</returns>
    public static IReadOnlyList<string> InsertHeaderColumns(
        object sheetObject,
        int headerRow,
        IReadOnlyList<CreatedColumn> expected,
        int firstColumn = 1)
    {
        var created = new List<string>();
        var cursor = firstColumn;

        for (var i = 0; i < expected.Count; i++)
        {
            var headers = ReadHeaderKeys(sheetObject, headerRow, firstColumn);

            var found = FindHeader(headers, expected[i].Keys, firstColumn);
            if (found > 0)
            {
                cursor = found + 1;
                continue;
            }

            // Ближайшая следующая колонка списка - ориентир: перед ней место пропущенной.
            var anchor = 0;
            for (var next = i + 1; next < expected.Count && anchor == 0; next++)
            {
                anchor = FindHeader(headers, expected[next].Keys, firstColumn);
            }

            int target;
            var inserted = false;
            if (IsHeaderEmpty(headers, cursor) && (anchor == 0 || cursor < anchor))
            {
                // Колонка на листе есть, у неё только нет названия: такую подписываем,
                // а не вставляем рядом ещё одну - иначе данные остались бы без заголовка.
                target = cursor;
            }
            else if (anchor > 0)
            {
                target = anchor;
                InsertColumns(sheetObject, anchor, 1);
                inserted = true;
            }
            else
            {
                target = Math.Max(LastHeaderColumn(headers), firstColumn - 1) + 1;
            }

            CopyHeaderFormat(sheetObject, headerRow, target, headers, inserted);
            SetValue(sheetObject, headerRow, target, expected[i].Title);
            created.Add(expected[i].Title);
            cursor = target + 1;
        }

        return created;
    }

    /// <summary>Нормализованные заголовки строки: индекс в списке - номер колонки минус один.</summary>
    private static IReadOnlyList<string> ReadHeaderKeys(object sheetObject, int headerRow, int firstColumn)
    {
        var bounds = GetUsedBounds(sheetObject);
        var lastColumn = Math.Max(bounds.FirstColumn + bounds.ColumnCount - 1, firstColumn);
        var grid = ReadBlock(sheetObject, headerRow, headerRow, 1, lastColumn, withFormulas: false);

        var keys = new string[lastColumn];
        for (var column = 1; column <= lastColumn; column++)
        {
            keys[column - 1] = TextUtils.NormalizeKey(grid.Text(headerRow, column));
        }

        return keys;
    }

    private static int FindHeader(IReadOnlyList<string> headers, IReadOnlyList<string> keys, int firstColumn)
    {
        for (var column = firstColumn; column <= headers.Count; column++)
        {
            if (keys.Contains(headers[column - 1], StringComparer.Ordinal))
            {
                return column;
            }
        }

        return 0;
    }

    /// <summary>Колонка есть, но заголовка у неё нет. За последним заголовком - тоже пусто.</summary>
    private static bool IsHeaderEmpty(IReadOnlyList<string> headers, int column) =>
        column > headers.Count || headers[column - 1].Length == 0;

    private static int LastHeaderColumn(IReadOnlyList<string> headers)
    {
        for (var column = headers.Count; column >= 1; column--)
        {
            if (headers[column - 1].Length > 0)
            {
                return column;
            }
        }

        return 0;
    }

    /// <summary>
    /// Оформление нового заголовка берётся у соседнего: сначала слева - слева колонки
    /// от вставки не сдвинулись. Если слева заголовков нет, берётся правый сосед: после
    /// вставки на его месте стоит тот заголовок, перед которым вставляли.
    /// </summary>
    private static void CopyHeaderFormat(
        object sheetObject, int headerRow, int column, IReadOnlyList<string> headers, bool inserted)
    {
        var source = 0;
        for (var left = column - 1; left >= 1 && source == 0; left--)
        {
            if (left <= headers.Count && headers[left - 1].Length > 0)
            {
                source = left;
            }
        }

        if (source == 0 && inserted)
        {
            source = column + 1;
        }

        if (source == 0)
        {
            return;
        }

        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic from = scope.Track(sheet.Cells[headerRow, source]);
        dynamic to = scope.Track(sheet.Cells[headerRow, column]);
        from.Copy(to);
    }

    /// <summary>
    /// Вставляет колонки перед указанной. Вставка именно внутрь занятого диапазона -
    /// это то, ради чего метод существует: так Excel сам растягивает формулы вида
    /// СУММ(N21:Y21) и СУММЕСЛИ($N$27:$Y$27; …), которые охватывают все блоки сразу.
    /// </summary>
    public static void InsertColumns(object sheetObject, int beforeColumn, int count)
    {
        if (count <= 0)
        {
            return;
        }

        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic columns = scope.Track(sheet.Range[
            ExcelColumn.ToLetters(beforeColumn) + ":" + ExcelColumn.ToLetters(beforeColumn + count - 1)]);
        columns.Insert(ExcelConstants.XlToRight);
    }

    public static void DeleteColumns(object sheetObject, int firstColumn, int count)
    {
        if (count <= 0)
        {
            return;
        }

        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic columns = scope.Track(sheet.Range[
            ExcelColumn.ToLetters(firstColumn) + ":" + ExcelColumn.ToLetters(firstColumn + count - 1)]);
        columns.Delete();
    }

    /// <summary>Копирует прямоугольник и вставляет его в другое место того же листа.</summary>
    public static void CopyRange(
        object applicationObject,
        object sheetObject,
        int firstRow,
        int firstColumn,
        int lastRow,
        int lastColumn,
        int targetRow,
        int targetColumn)
    {
        dynamic application = applicationObject;
        dynamic sheet = sheetObject;

        using (var scope = new ComScope())
        {
            dynamic source = scope.Track(sheet.Range[BuildReference(firstRow, firstColumn, lastRow, lastColumn)]);
            dynamic destination = scope.Track(sheet.Range[BuildReference(
                targetRow,
                targetColumn,
                targetRow + (lastRow - firstRow),
                targetColumn + (lastColumn - firstColumn))]);
            source.Copy(destination);
        }

        application.CutCopyMode = false;
    }

    /// <summary>Сортирует строки прямоугольника по одной колонке по возрастанию.</summary>
    public static void SortRange(
        object sheetObject,
        int firstRow,
        int lastRow,
        int firstColumn,
        int lastColumn,
        int keyColumn)
    {
        if (lastRow <= firstRow)
        {
            return;
        }

        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic range = scope.Track(sheet.Range[BuildReference(firstRow, firstColumn, lastRow, lastColumn)]);
        dynamic key = scope.Track(sheet.Range[BuildReference(firstRow, keyColumn, lastRow, keyColumn)]);

        // Header:=xlNo - строка заголовков в диапазон не входит, сортируются только данные.
        range.Sort(key, ExcelConstants.XlAscending, Type.Missing, Type.Missing, ExcelConstants.XlAscending,
            Type.Missing, ExcelConstants.XlAscending, ExcelConstants.XlNo);
    }

    private static string BuildReference(int firstRow, int firstColumn, int lastRow, int lastColumn) =>
        ExcelColumn.ToLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture) + ":" +
        ExcelColumn.ToLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture);

    /// <summary>Очищает значения в прямоугольнике, оставляя оформление.</summary>
    public static void ClearBlock(object sheetObject, int firstRow, int lastRow, int column)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();

        var letters = ExcelColumn.ToLetters(column);
        var reference =
            letters + firstRow.ToString(CultureInfo.InvariantCulture) + ":" +
            letters + lastRow.ToString(CultureInfo.InvariantCulture);

        dynamic range = scope.Track(sheet.Range[reference]);
        range.ClearContents();
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

    /// <summary>Одна высота для строк подряд - одним обменом с COM.</summary>
    public static void SetRowsHeight(object sheetObject, int firstRow, int lastRow, double height)
    {
        if (lastRow < firstRow)
        {
            return;
        }

        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic rows = scope.Track(sheet.Rows[RowReference(firstRow, lastRow)]);
        rows.RowHeight = height;
    }

    /// <summary>Формулы строки в относительной записи R1C1. Ячейка без формулы - null.</summary>
    public static string?[] ReadRowFormulasR1C1(object sheetObject, int row, int firstColumn, int lastColumn)
    {
        dynamic sheet = sheetObject;
        using var scope = new ComScope();
        dynamic range = scope.Track(sheet.Range[BuildRangeReference(row, firstColumn, lastColumn)]);
        object? raw = range.FormulaR1C1;

        var count = lastColumn - firstColumn + 1;
        var result = new string?[count];
        for (var i = 0; i < count; i++)
        {
            var text = raw is object[,] array
                ? array[array.GetLowerBound(0), array.GetLowerBound(1) + i] as string
                : count == 1 ? raw as string : null;
            result[i] = text is not null && text.StartsWith("=", StringComparison.Ordinal) ? text : null;
        }

        return result;
    }

    /// <summary>
    /// Переносит оформление строки-образца (заливку, шрифт, границы, формат чисел, условное
    /// форматирование) на строки ниже, не трогая их значений и формул.
    /// </summary>
    public static void CopyRowFormats(
        object applicationObject,
        object sheetObject,
        int templateRow,
        int firstRow,
        int lastRow,
        ColumnRange columns)
    {
        if (lastRow < firstRow)
        {
            return;
        }

        dynamic application = applicationObject;
        dynamic sheet = sheetObject;

        using (var scope = new ComScope())
        {
            dynamic source = scope.Track(sheet.Range[BlockReference(templateRow, templateRow, columns)]);
            source.Copy();
            dynamic destination = scope.Track(sheet.Range[BlockReference(firstRow, lastRow, columns)]);
            destination.PasteSpecial(ExcelConstants.XlPasteFormats);
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
