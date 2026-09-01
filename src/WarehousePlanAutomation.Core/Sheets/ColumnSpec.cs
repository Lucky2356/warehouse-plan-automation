namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Описание искомой колонки: отображаемое имя (для сообщений об ошибке) и допустимые написания
/// заголовка. Буквы колонок нигде не фиксируются - колонка ищется по заголовку.
/// </summary>
public sealed class ColumnSpec
{
    public ColumnSpec(string displayName, string[] aliases, bool exactOnly = false)
    {
        DisplayName = displayName;
        Aliases = aliases;
        ExactOnly = exactOnly;
    }

    public string DisplayName { get; }

    /// <summary>Нормализованные (в нижнем регистре) варианты заголовка.</summary>
    public string[] Aliases { get; }

    /// <summary>
    /// Если true, заголовок сопоставляется только полным совпадением.
    /// Нужно для колонки «%» журнала, рядом с которой есть «% оклейка», «% отгрузка», «% сборка».
    /// </summary>
    public bool ExactOnly { get; }
}
