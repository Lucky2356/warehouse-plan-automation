namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Описание искомой колонки: отображаемое имя (для сообщений об ошибке) и допустимые написания
/// заголовка. Буквы колонок нигде не фиксируются - колонка ищется по заголовку.
/// </summary>
public sealed class ColumnSpec
{
    public ColumnSpec(string displayName, string[] aliases, bool exactOnly = false, bool prefixOnly = false)
    {
        DisplayName = displayName;
        Aliases = aliases;
        ExactOnly = exactOnly;
        PrefixOnly = prefixOnly;
    }

    public string DisplayName { get; }

    /// <summary>Нормализованные (в нижнем регистре) варианты заголовка.</summary>
    public string[] Aliases { get; }

    /// <summary>
    /// Если true, заголовок сопоставляется только полным совпадением.
    /// Нужно для колонки «%» журнала, рядом с которой есть «% оклейка», «% отгрузка», «% сборка».
    /// </summary>
    public bool ExactOnly { get; }

    /// <summary>
    /// Если true, заголовок может продолжаться после названия, но начинаться должен с него:
    /// «в подтоварку Мск» - это «в подтоварку». Вхождение в середину («урезать в подтоварку»)
    /// не считается. Полное совпадение по-прежнему сильнее.
    /// </summary>
    public bool PrefixOnly { get; }
}
