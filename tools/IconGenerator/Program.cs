using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace IconGenerator;

/// <summary>
/// Рисует иконку приложения «лист плана» и собирает из неё .ico.
///
/// В крупных размерах это лист документа: белый бланк с акцентной шапкой и строками.
/// В мелких (32 px и меньше) контур и белая заливка сливаются в кашу, поэтому силуэт
/// заливается целиком, строк остаётся две и они белые - так знак читается даже в 16 px.
/// </summary>
internal static class Program
{
    /// <summary>Размеры кадров в .ico. Windows выбирает подходящий сам.</summary>
    private static readonly int[] Sizes = { 16, 20, 24, 32, 48, 64, 128, 256 };

    /// <summary>До этого размера включительно рисуется сплошной силуэт.</summary>
    private const int SolidUpTo = 32;

    private static readonly Color Accent = Color.FromRgb(0x4C, 0x4F, 0xA6);
    private static readonly Color AccentDark = Color.FromRgb(0x37, 0x39, 0x82);
    private static readonly Color Paper = Color.FromRgb(0xFF, 0xFF, 0xFF);
    private static readonly Color SolidLine = Color.FromRgb(0xFF, 0xFF, 0xFF);

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length != 1)
        {
            Console.Error.WriteLine(
                "Использование: dotnet run --project tools/IconGenerator -- <путь к app.ico>");
            return 1;
        }

        var outputPath = Path.GetFullPath(args[0]);
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var frames = Sizes.Select(RenderPng).ToList();
        WriteIcon(outputPath, Sizes, frames);

        Console.WriteLine(
            "Записано " + outputPath + ": кадров " + Sizes.Length +
            ", " + new FileInfo(outputPath).Length.ToString(CultureInfo.InvariantCulture) + " байт.");
        return 0;
    }

    private static byte[] RenderPng(int size)
    {
        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen())
        {
            if (size <= SolidUpTo)
            {
                DrawSolid(dc, size);
            }
            else
            {
                DrawSheet(dc, size);
            }
        }

        var bitmap = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(visual);

        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));

        using var stream = new MemoryStream();
        encoder.Save(stream);
        return stream.ToArray();
    }

    /// <summary>Крупный размер: белый лист с акцентной шапкой и строками.</summary>
    private static void DrawSheet(DrawingContext dc, int size)
    {
        var u = size / 64d;

        var sheet = new RectangleGeometry(
            new Rect(11 * u, 5 * u, 42 * u, 54 * u),
            5 * u,
            5 * u);

        var border = new Pen(new SolidColorBrush(Accent), 3.4 * u);
        dc.DrawGeometry(new SolidColorBrush(Paper), border, sheet);

        // Шапка листа: прямоугольник, обрезанный по скруглению самого листа.
        dc.PushClip(sheet);
        dc.DrawRectangle(
            new SolidColorBrush(Accent),
            null,
            new Rect(11 * u, 5 * u, 42 * u, 13 * u));
        dc.Pop();

        var line = new Pen(new SolidColorBrush(Accent), 3.2 * u) { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round };
        dc.DrawLine(line, new Point(19 * u, 28 * u), new Point(45 * u, 28 * u));
        dc.DrawLine(line, new Point(19 * u, 37 * u), new Point(45 * u, 37 * u));
        dc.DrawLine(line, new Point(19 * u, 46 * u), new Point(33 * u, 46 * u));
    }

    /// <summary>Мелкий размер: сплошной силуэт, строки светлые.</summary>
    private static void DrawSolid(DrawingContext dc, int size)
    {
        var u = size / 64d;

        var sheet = new RectangleGeometry(
            new Rect(9 * u, 4 * u, 46 * u, 56 * u),
            6 * u,
            6 * u);

        dc.DrawGeometry(new SolidColorBrush(Accent), null, sheet);

        dc.PushClip(sheet);
        dc.DrawRectangle(
            new SolidColorBrush(AccentDark),
            null,
            new Rect(9 * u, 4 * u, 46 * u, 14 * u));
        dc.Pop();

        var line = new Pen(new SolidColorBrush(SolidLine), 6 * u) { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round };
        dc.DrawLine(line, new Point(19 * u, 31 * u), new Point(45 * u, 31 * u));
        dc.DrawLine(line, new Point(19 * u, 44 * u), new Point(45 * u, 44 * u));
    }

    /// <summary>
    /// Собирает .ico из PNG-кадров. Windows Vista и новее читают PNG внутри .ico
    /// напрямую, поэтому перекодировать в BMP не нужно.
    /// </summary>
    private static void WriteIcon(string path, IReadOnlyList<int> sizes, IReadOnlyList<byte[]> frames)
    {
        using var file = File.Create(path);
        using var writer = new BinaryWriter(file);

        writer.Write((ushort)0);            // зарезервировано
        writer.Write((ushort)1);            // тип: иконка
        writer.Write((ushort)frames.Count);

        var offset = 6 + (16 * frames.Count);
        for (var i = 0; i < frames.Count; i++)
        {
            // 256 записывается нулём: в поле помещается только один байт.
            writer.Write((byte)(sizes[i] >= 256 ? 0 : sizes[i]));
            writer.Write((byte)(sizes[i] >= 256 ? 0 : sizes[i]));
            writer.Write((byte)0);          // палитра не используется
            writer.Write((byte)0);          // зарезервировано
            writer.Write((ushort)1);        // плоскостей
            writer.Write((ushort)32);       // бит на пиксель
            writer.Write(frames[i].Length);
            writer.Write(offset);
            offset += frames[i].Length;
        }

        foreach (var frame in frames)
        {
            writer.Write(frame);
        }
    }
}
