using System.Runtime.InteropServices;
using WarehousePlanAutomation.Core.Logging;

namespace WarehousePlanAutomation.Excel;

/// <summary>Прямоугольник листа в строках и колонках.</summary>
internal readonly record struct SheetArea(int FirstRow, int LastRow, int FirstColumn, int LastColumn)
{
    public bool Contains(int row, int column) =>
        row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
}

/// <summary>
/// Картинки на листе.
///
/// В инвойсе у каждой строки товара своя фотография, и она должна доехать до листа «Цены»,
/// а оттуда - до заголовка своего блока на листе «Распред». Сами по себе картинки туда
/// не попадают: копирование строки или блока размножает ту, что была в образце, поэтому
/// перед вставкой старые картинки из области удаляются.
///
/// Excel умеет переносить картинку между листами только через буфер обмена, а вставляет
/// её туда, где стоит выделение. Поэтому лист приходится делать активным - на невидимом
/// приложении это ничего не показывает.
/// </summary>
internal static class ExcelPictures
{
    /// <summary>Тип фигуры «картинка» (msoPicture).</summary>
    private const int MsoPicture = 13;

    /// <summary>
    /// Картинка, связанная с файлом (msoLinkedPicture). Так ложатся фотографии, которые
    /// вставлены макросом по ссылке, - для программы это такая же фотография товара.
    /// </summary>
    private const int MsoLinkedPicture = 11;

    private static bool IsPicture(dynamic shape)
    {
        int type = shape.Type;
        return type is MsoPicture or MsoLinkedPicture;
    }

    /// <summary>Картинка двигается вместе с ячейками, но не растягивается (xlMove).</summary>
    private const int XlMove = 2;

    /// <summary>
    /// Картинки прямоугольника: строка их привязки - ключ, значение - сама фигура.
    /// Если в строке несколько картинок, берётся первая.
    /// </summary>
    public static IReadOnlyDictionary<int, object> ByRow(object sheetObject, SheetArea area, ComScope scope)
    {
        dynamic sheet = sheetObject;
        var result = new Dictionary<int, object>();

        dynamic shapes = scope.Track(sheet.Shapes);
        int count = shapes.Count;

        for (var index = 1; index <= count; index++)
        {
            dynamic shape = scope.Track(shapes.Item(index));
            if (!IsPicture(shape))
            {
                continue;
            }

            dynamic cell = scope.Track(shape.TopLeftCell);
            int row = cell.Row;
            int column = cell.Column;

            if (area.Contains(row, column) && !result.ContainsKey(row))
            {
                result[row] = (object)shape;
            }
        }

        return result;
    }

    /// <summary>Удаляет картинки, привязанные внутри прямоугольника. Возвращает, сколько удалено.</summary>
    /// <param name="overlapping">
    /// Удалять и картинки, которые только заходят в прямоугольник, а привязаны за ним:
    /// в заголовках блоков «Распреда» фотография, сдвинутая на строку-другую, всё равно
    /// лежит под новой и должна уйти.
    /// </param>
    public static int DeleteIn(object sheetObject, SheetArea area, bool overlapping = false)
    {
        if (overlapping)
        {
            return DeleteOverlapping(sheetObject, area);
        }

        dynamic sheet = sheetObject;
        var deleted = 0;

        using var scope = new ComScope();
        dynamic shapes = scope.Track(sheet.Shapes);
        int count = shapes.Count;

        // С конца: удаление сдвигает номера в коллекции.
        for (var index = count; index >= 1; index--)
        {
            dynamic shape = scope.Track(shapes.Item(index));
            if (!IsPicture(shape))
            {
                continue;
            }

            dynamic cell = scope.Track(shape.TopLeftCell);
            if (!area.Contains((int)cell.Row, (int)cell.Column))
            {
                continue;
            }

            shape.Delete();
            deleted++;
        }

        return deleted;
    }

    private static int DeleteOverlapping(object sheetObject, SheetArea area)
    {
        dynamic sheet = sheetObject;
        var deleted = 0;

        using var scope = new ComScope();
        dynamic shapes = scope.Track(sheet.Shapes);
        int count = shapes.Count;

        for (var index = count; index >= 1; index--)
        {
            dynamic shape = scope.Track(shapes.Item(index));
            if (!IsPicture(shape))
            {
                continue;
            }

            dynamic topLeft = scope.Track(shape.TopLeftCell);
            dynamic bottomRight = scope.Track(shape.BottomRightCell);
            int top = topLeft.Row;
            int left = topLeft.Column;
            int bottom = bottomRight.Row;
            int right = bottomRight.Column;

            var intersects = top <= area.LastRow && bottom >= area.FirstRow &&
                             left <= area.LastColumn && right >= area.FirstColumn;
            if (!intersects)
            {
                continue;
            }

            shape.Delete();
            deleted++;
        }

        return deleted;
    }

    /// <summary>
    /// Копирует картинку на другой лист и ставит её в ячейку, вписывая в заданную рамку
    /// с сохранением пропорций. Возвращает false, если Excel вставку не выполнил.
    ///
    /// Буфер обмена общий на весь компьютер, и на рабочем месте им пользуются все: открытый
    /// рядом Excel, мессенджеры, удалённый рабочий стол. Если копирование не дошло до буфера,
    /// вставка берёт то, что в нём лежало раньше, - и во всех блоках оказывается одна и та же
    /// чужая фотография. Поэтому буфер перед копированием очищается, после копирования
    /// проверяется, что в нём наше, а вставленная картинка сверяется по размеру с исходной.
    /// </summary>
    /// <param name="bottomRight">
    /// Прижать картинку к правому нижнему углу ячейки и не выпускать за её ширину:
    /// так текст ячейки, который стоит сверху слева, остаётся виден.
    /// </param>
    public static bool CopyTo(
        object applicationObject,
        object shapeObject,
        object targetSheetObject,
        int row,
        int column,
        double maxWidth,
        double maxHeight,
        bool fillFrame,
        IAppLogger logger,
        bool bottomRight = false)
    {
        dynamic application = applicationObject;
        dynamic shape = shapeObject;
        dynamic target = targetSheetObject;

        double sourceWidth = shape.Width;
        double sourceHeight = shape.Height;

        // Если буфер в эту долю секунды занял другой процесс, перенос не удаётся.
        // Такая ошибка проходит сама, поэтому перенос повторяется.
        for (var attempt = 1; attempt <= PasteAttempts; attempt++)
        {
            try
            {
                using var scope = new ComScope();
                dynamic shapes = scope.Track(target.Shapes);
                int before = shapes.Count;

                Clipboard.Clear();
                var cleared = Clipboard.Sequence();
                shape.Copy();
                var copied = Clipboard.Sequence();
                if (copied == cleared)
                {
                    logger.Warning("Картинка для строки " + row + " не скопировалась в буфер, попытка " + attempt + ".");
                    Thread.Sleep(PasteRetryDelayMs);
                    continue;
                }

                target.Activate();
                dynamic anchor = scope.Track(target.Cells[row, column]);
                anchor.Select();

                if (Clipboard.Sequence() != copied)
                {
                    logger.Warning("Буфер обмена перехватил другой процесс, строка " + row + ", попытка " + attempt + ".");
                    Thread.Sleep(PasteRetryDelayMs);
                    continue;
                }

                target.Paste();
                application.CutCopyMode = false;

                if ((int)shapes.Count <= before)
                {
                    logger.Warning("Картинка в строку " + row + " не вставилась, попытка " + attempt + ".");
                    Thread.Sleep(PasteRetryDelayMs);
                    continue;
                }

                dynamic pasted = scope.Track(shapes.Item(shapes.Count));

                // Вставленная копия выходит того же размера, что исходная картинка.
                // Другой размер значит, что вставилось чужое содержимое буфера.
                double pastedWidth = pasted.Width;
                double pastedHeight = pasted.Height;
                if (!SameSize(sourceWidth, sourceHeight, pastedWidth, pastedHeight))
                {
                    logger.Warning(
                        "В строку " + row + " вставилась не та картинка (" + pastedWidth.ToString("0.#") + "x" +
                        pastedHeight.ToString("0.#") + " вместо " + sourceWidth.ToString("0.#") + "x" +
                        sourceHeight.ToString("0.#") + "), попытка " + attempt + ".");
                    pasted.Delete();
                    Thread.Sleep(PasteRetryDelayMs);
                    continue;
                }
                double cellWidth = anchor.Width;
                double cellHeight = anchor.Height;
                Fit(pasted, bottomRight ? Math.Min(maxWidth, cellWidth) : maxWidth, maxHeight, fillFrame);
                pasted.Placement = XlMove;

                if (bottomRight)
                {
                    pasted.Left = (double)anchor.Left + Math.Max(0d, cellWidth - (double)pasted.Width);
                    pasted.Top = (double)anchor.Top + Math.Max(0d, cellHeight - (double)pasted.Height);
                }
                else
                {
                    pasted.Left = (double)anchor.Left;
                    pasted.Top = (double)anchor.Top;
                }

                return true;
            }
            catch (COMException ex)
            {
                logger.Warning("Не удалось перенести картинку в строку " + row + ", попытка " + attempt + ".", ex);
                Thread.Sleep(PasteRetryDelayMs);
            }
        }

        return false;
    }

    private const int PasteAttempts = 5;

    /// <summary>Размеры совпадают с точностью до округления Excel: 2 % или 1 пункт.</summary>
    private static bool SameSize(double width, double height, double otherWidth, double otherHeight)
    {
        static bool Near(double a, double b) => Math.Abs(a - b) <= Math.Max(1d, Math.Abs(a) * 0.02);
        return Near(width, otherWidth) && Near(height, otherHeight);
    }

    /// <summary>
    /// Буфер обмена Windows напрямую: очистить и узнать номер его содержимого.
    /// Номер меняется при каждой записи в буфер - по нему видно, дошло ли копирование.
    /// </summary>
    private static class Clipboard
    {
        [DllImport("user32.dll")]
        private static extern uint GetClipboardSequenceNumber();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool OpenClipboard(IntPtr owner);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool EmptyClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseClipboard();

        public static uint Sequence() => GetClipboardSequenceNumber();

        /// <summary>Очищает буфер. Если он занят, не страшно: проверка номера всё равно поймает сбой.</summary>
        public static void Clear()
        {
            for (var attempt = 0; attempt < 5; attempt++)
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    try
                    {
                        EmptyClipboard();
                    }
                    finally
                    {
                        CloseClipboard();
                    }

                    return;
                }

                Thread.Sleep(50);
            }
        }
    }

    private const int PasteRetryDelayMs = 300;

    /// <summary>msoTrue для ScaleWidth/ScaleHeight: масштаб считается от исходного размера картинки.</summary>
    private const int MsoTrue = -1;

    private const int MsoFalse = 0;

    /// <summary>
    /// Вписывает картинку в рамку, сохраняя пропорции. При <paramref name="fillFrame"/>
    /// картинка растягивается до рамки, иначе только уменьшается, если не помещается.
    ///
    /// В инвойсе картинки бывают растянуты или сжаты по одной стороне - такая копия
    /// выходила узкой полоской. Поэтому сначала возвращаются исходные пропорции снимка,
    /// и только потом он вписывается в рамку: ширина и высота задаются обе, явно.
    /// </summary>
    private static void Fit(dynamic shape, double maxWidth, double maxHeight, bool fillFrame)
    {
        try
        {
            shape.LockAspectRatio = MsoFalse;
            shape.ScaleHeight(1f, MsoTrue);
            shape.ScaleWidth(1f, MsoTrue);
        }
        catch (COMException)
        {
            // У фигуры нет исходного размера (не снимок) - пропорции остаются текущими.
        }

        double width = shape.Width;
        double height = shape.Height;

        if (width <= 0 || height <= 0 || maxWidth <= 0 || maxHeight <= 0)
        {
            return;
        }

        var scale = Math.Min(maxWidth / width, maxHeight / height);
        if (!fillFrame && scale >= 1d)
        {
            scale = 1d;
        }

        shape.Width = width * scale;
        shape.Height = height * scale;
        shape.LockAspectRatio = MsoTrue;
    }
}
