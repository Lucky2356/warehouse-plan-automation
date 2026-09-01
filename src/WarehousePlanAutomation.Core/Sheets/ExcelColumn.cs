using System.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>Преобразование номера колонки Excel в буквенное обозначение и обратно.</summary>
public static class ExcelColumn
{
    public static string ToLetters(int column)
    {
        if (column < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(column));
        }

        var builder = new StringBuilder();
        var value = column;
        while (value > 0)
        {
            var remainder = (value - 1) % 26;
            builder.Insert(0, (char)('A' + remainder));
            value = (value - 1) / 26;
        }

        return builder.ToString();
    }

    public static int FromLetters(string letters)
    {
        if (string.IsNullOrWhiteSpace(letters))
        {
            throw new ArgumentException("Пустое обозначение колонки.", nameof(letters));
        }

        var result = 0;
        foreach (var ch in letters.Trim().ToUpperInvariant())
        {
            if (ch < 'A' || ch > 'Z')
            {
                throw new ArgumentException("Некорректное обозначение колонки: " + letters, nameof(letters));
            }

            result = (result * 26) + (ch - 'A' + 1);
        }

        return result;
    }
}
