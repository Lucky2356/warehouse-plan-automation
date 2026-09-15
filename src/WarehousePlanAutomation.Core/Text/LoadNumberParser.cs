using System.Globalization;

namespace WarehousePlanAutomation.Core.Text;

/// <summary>
/// Извлечение номера загрузки из комментария и текста «Поставки» до слов «Номер загрузки».
/// </summary>
public static class LoadNumberParser
{
    public const string Marker = "Номер загрузки";

    /// <summary>Сколько разделителей допускается между словами «Номер загрузки» и самим номером.</summary>
    private const int MaxSeparatorLength = 5;

    public static bool TryExtract(string? comment, out long loadNumber)
    {
        loadNumber = 0;
        var normalized = TextUtils.Normalize(comment);
        var markerIndex = normalized.IndexOf(Marker, StringComparison.OrdinalIgnoreCase);
        if (markerIndex < 0)
        {
            return false;
        }

        // Между словами «Номер загрузки» и самим номером допускаются только разделители
        // (пробел, двоеточие, знак номера) и не более MaxSeparatorLength подряд. Иначе
        // в комментарии вида «Номер загрузки не указан, отгрузка 05.09» номером загрузки
        // ошибочно считалось бы первое попавшееся дальше число.
        var index = markerIndex + Marker.Length;
        var limit = Math.Min(normalized.Length, index + MaxSeparatorLength);
        while (index < limit && !char.IsDigit(normalized[index]))
        {
            if (char.IsLetter(normalized[index]))
            {
                return false;
            }

            index++;
        }

        if (index >= normalized.Length || !char.IsDigit(normalized[index]))
        {
            return false;
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
