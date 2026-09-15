namespace WarehousePlanAutomation.Core.Text;

/// <summary>
/// Ошибки формул, как их отдаёт Excel через Value2: отрицательным целым кодом,
/// а не текстом «#Н/Д». Без этой проверки код ошибки прочитался бы как обычное
/// число вроде -2146826246 и попал бы в расчёты.
/// </summary>
public static class CellError
{
    public const int NotAvailable = -2146826246;
    public const int DivideByZero = -2146826281;
    public const int Value = -2146826273;
    public const int Name = -2146826259;
    public const int Reference = -2146826265;
    public const int Number = -2146826252;
    public const int Null = -2146826288;

    private static readonly int[] Codes =
    {
        NotAvailable, DivideByZero, Value, Name, Reference, Number, Null,
    };

    public static bool IsError(object? value) => value switch
    {
        int code => Array.IndexOf(Codes, code) >= 0,
        double number => number <= int.MaxValue && number >= int.MinValue &&
                         Math.Abs(number - Math.Round(number)) < 1e-9 &&
                         Array.IndexOf(Codes, (int)Math.Round(number)) >= 0,
        string text => text.StartsWith("#", StringComparison.Ordinal),
        _ => false,
    };

    /// <summary>
    /// Может ли Excel принять эту строку за число при записи.
    ///
    /// Значение, записанное в ячейку, Excel разбирает так же, как набранное руками,
    /// и запятые считает разделителями разрядов. Из-за этого перечень магазинов
    /// «150,155,156,158,…» превращается в 1,5E+113. Такие строки нужно писать
    /// в текстовом формате.
    /// </summary>
    public static bool LooksNumericToExcel(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return false;
        }

        var hasDigit = false;
        foreach (var ch in value)
        {
            if (char.IsDigit(ch))
            {
                hasDigit = true;
                continue;
            }

            if (ch is ',' or '.' or ' ' or '+' or '-' or 'e' or 'E' or ' ' or '%')
            {
                continue;
            }

            return false;
        }

        return hasDigit;
    }

    public static bool IsNotAvailable(object? value) => value switch
    {
        int code => code == NotAvailable,
        double number => Math.Abs(number - NotAvailable) < 0.5,
        string text => TextUtils.EqualsKey(text, "#н/д") || TextUtils.EqualsKey(text, "#n/a"),
        _ => false,
    };

    /// <summary>
    /// «#ССЫЛКА!» - ссылка не разрешилась. У колонок со стенками это значит не «позиции
    /// нет в стенках», а «файл стенок не прочитался»: он лежит на сетевом диске, и без
    /// доступа к нему Excel теряет сохранённые значения при первом же пересчёте.
    /// Причина другая, поэтому и сообщение другое.
    /// </summary>
    public static bool IsBrokenReference(object? value) => value switch
    {
        int code => code == Reference,
        double number => Math.Abs(number - Reference) < 0.5,
        string text => TextUtils.EqualsKey(text, "#ссылка!") || TextUtils.EqualsKey(text, "#ref!"),
        _ => false,
    };
}
