namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Числовые значения перечислений Excel. Используется позднее связывание (late binding),
/// поэтому библиотеки типов Office на машине сборки не требуются.
/// </summary>
internal static class ExcelConstants
{
    public const int XlCalculationAutomatic = -4105;
    public const int XlCalculationManual = -4135;
    public const int XlDown = -4121;
    public const int XlUp = -4162;
    public const int XlNoChange = -4142;
    public const int XlToRight = -4161;

    /// <summary>Range.PasteSpecial: вставить только форматы.</summary>
    public const int XlPasteFormats = -4122;

    /// <summary>Range.Sort: порядок по возрастанию и «в диапазоне нет строки заголовков».</summary>
    public const int XlAscending = 1;

    public const int XlNo = 2;

    /// <summary>Interior.ColorIndex: «без заливки».</summary>
    public const int XlColorIndexNone = -4142;

    /// <summary>Font.ColorIndex: «авто».</summary>
    public const int XlColorIndexAutomatic = -4105;

    /// <summary>Индекс параметра «десятичный разделитель» в Application.International.</summary>
    public const int XlDecimalSeparator = 3;

    /// <summary>Индекс параметра «разделитель списка» в Application.International.</summary>
    public const int XlListSeparator = 5;

    /// <summary>Размер листа Excel: нужен, когда прямоугольник задаётся «весь лист».</summary>
    public const int LastRow = 1048576;

    public const int LastColumn = 16384;
}
