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
            if ((int)shape.Type != MsoPicture)
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
    public static int DeleteIn(object sheetObject, SheetArea area)
    {
        dynamic sheet = sheetObject;
        var deleted = 0;

        using var scope = new ComScope();
        dynamic shapes = scope.Track(sheet.Shapes);
        int count = shapes.Count;

        // С конца: удаление сдвигает номера в коллекции.
        for (var index = count; index >= 1; index--)
        {
            dynamic shape = scope.Track(shapes.Item(index));
            if ((int)shape.Type != MsoPicture)
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

    /// <summary>
    /// Копирует картинку на другой лист и ставит её в ячейку, вписывая в заданную рамку
    /// с сохранением пропорций. Возвращает false, если Excel вставку не выполнил.
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

        // Буфер обмена в Windows общий: если его в эту долю секунды занял другой процесс,
        // вставка не удаётся. Такая ошибка проходит сама, поэтому перенос повторяется.
        for (var attempt = 1; attempt <= PasteAttempts; attempt++)
        {
            try
            {
                using var scope = new ComScope();
                dynamic shapes = scope.Track(target.Shapes);
                int before = shapes.Count;

                shape.Copy();
                target.Activate();
                dynamic anchor = scope.Track(target.Cells[row, column]);
                anchor.Select();
                target.Paste();
                application.CutCopyMode = false;

                if ((int)shapes.Count <= before)
                {
                    logger.Warning("Картинка в строку " + row + " не вставилась, попытка " + attempt + ".");
                    Thread.Sleep(PasteRetryDelayMs);
                    continue;
                }

                dynamic pasted = scope.Track(shapes.Item(shapes.Count));
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

    private const int PasteAttempts = 3;

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
