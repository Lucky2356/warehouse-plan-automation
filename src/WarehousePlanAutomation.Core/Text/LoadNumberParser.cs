using System.Globalization;

namespace WarehousePlanAutomation.Core.Text;

/// <summary>
/// Извлечение номера загрузки из комментария и текста «Поставки» до слов «Номер загрузки».
/// </summary>
public static class LoadNumberParser
{
    public const string Marker = "Номер загрузки";

    public static bool TryExtract(string? comment, out long loadNumber)
    {
        loadNumber = 0;
        var normalized = TextUtils.Normalize(comment);
        var markerIndex = normalized.IndexOf(Marker, StringComparison.OrdinalIgnoreCase);
        if (markerIndex < 0)
        {
            return false;
        }

        var index = markerIndex + Marker.Length;
        while (index < normalized.Length && !char.IsDigit(normalized[index]))
        {
            index++;
        }

        var start = index;
        while (index < normalized.Length && char.IsDigit(normalized[index]))
        {
            index++;
        }

        if (index == start)
        {
            return false;
        }

        return long.TryParse(
            normalized.AsSpan(start, index - start),
            NumberStyles.None,
            CultureInfo.InvariantCulture,
            out loadNumber);
    }

    /// <summary>
    /// Текст для колонки «Поставки»: часть комментария от начала и до слов «Номер загрузки».
    /// Сами слова и всё, что после них, исключаются.
    /// </summary>
    public static string ExtractSuppliesText(string? comment)
    {
        var normalized = TextUtils.Normalize(comment);
        var markerIndex = normalized.IndexOf(Marker, StringComparison.OrdinalIgnoreCase);
        return markerIndex < 0 ? normalized : normalized[..markerIndex].Trim();
    }
}
