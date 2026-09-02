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

    /// <summary>Interior.ColorIndex: «без заливки».</summary>
    public const int XlColorIndexNone = -4142;

    /// <summary>Индекс параметра «разделитель списка» в Application.International.</summary>
    public const int XlListSeparator = 5;
}
