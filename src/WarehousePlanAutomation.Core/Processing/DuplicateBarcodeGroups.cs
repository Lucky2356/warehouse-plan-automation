namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Раскраска одинаковых штрихкодов листа «Цены»: у каждой группы повторов свой цвет,
/// чтобы с одного взгляда было видно, какие строки - пара. Цвета светлые, текст на них
/// читается, и ни один не совпадает с розовой «Плохо» и зелёной «Хорошо» пометками.
/// Групп больше, чем цветов, - цвета идут по кругу.
/// </summary>
public static class DuplicateBarcodeGroups
{
    /// <summary>
    /// Заливки в порядке BGR, как их принимает Excel. Повторный запуск снимает их как свои,
    /// поэтому оттенки взяты чуть в стороне от стандартной палитры Excel: цвет, которым
    /// аналитик покрасила ячейку сама, с ними не совпадёт.
    /// </summary>
    public static IReadOnlyList<int> Palette { get; } = new[]
    {
        Bgr(0xB9, 0xD5, 0xF1), // голубой
        Bgr(0xFF, 0xE3, 0x8F), // жёлтый
        Bgr(0xDA, 0xC1, 0xEB), // сиреневый
        Bgr(0xF9, 0xC9, 0xA9), // персиковый
        Bgr(0xB3, 0xE3, 0xE7), // бирюзовый
        Bgr(0xE8, 0xD2, 0xAF), // бежевый
        Bgr(0xC7, 0xC7, 0xF7), // лавандовый
        Bgr(0xCF, 0xCB, 0xCB), // серый
    };

    /// <summary>
    /// Номер группы для каждой строки: строки с одинаковым штрихкодом получают один номер,
    /// группы нумеруются по первому появлению. У строки без повтора и без штрихкода - null.
    /// </summary>
    public static IReadOnlyList<int?> Assign(IReadOnlyList<string> barcodes)
    {
        var counts = barcodes
            .Where(barcode => barcode.Length > 0)
            .GroupBy(barcode => barcode, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

        var groups = new Dictionary<string, int>(StringComparer.Ordinal);
        var result = new int?[barcodes.Count];
        for (var i = 0; i < barcodes.Count; i++)
        {
            var barcode = barcodes[i];
            if (barcode.Length == 0 || counts[barcode] < 2)
            {
                continue;
            }

            if (!groups.TryGetValue(barcode, out var group))
            {
                group = groups.Count;
                groups[barcode] = group;
            }

            result[i] = group;
        }

        return result;
    }

    /// <summary>Цвет группы; групп больше, чем цветов, - цвета повторяются по кругу.</summary>
    public static int ColorOf(int group) => Palette[group % Palette.Count];

    private static int Bgr(int red, int green, int blue) => (blue << 16) | (green << 8) | red;
}
