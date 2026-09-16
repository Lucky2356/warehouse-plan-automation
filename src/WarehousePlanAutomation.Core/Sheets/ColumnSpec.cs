namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Описание искомой колонки: отображаемое имя (для сообщений об ошибке) и допустимые написания
/// заголовка. Буквы колонок нигде не фиксируются - колонка ищется по заголовку.
/// </summary>
public sealed class ColumnSpec
{
    public ColumnSpec(
        string displayName,
        string[] aliases,
        bool exactOnly = false,
        bool prefixOnly = false,
        bool optional = false)
    {
        DisplayName = displayName;
        Aliases = aliases;
        ExactOnly = exactOnly;
        PrefixOnly = prefixOnly;
        Optional = optional;
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

    /// <summary>
    /// Если true, без такой колонки книга всё равно разбирается: её либо дописывает сама
    /// программа, либо она нужна не всегда. На выбор строки заголовков не влияет.
    /// </summary>
    public bool Optional { get; }
}

/// <summary>
/// Колонка, которую программа дописывает в книгу, если её нет: под каким названием создать
/// и по каким заголовкам искать уже готовую.
/// </summary>
/// <param name="Name">Название колонки в разметке листа (<see cref="ColumnSpec.DisplayName"/>).</param>
/// <param name="Title">Заголовок, который пишется в новую колонку.</param>
/// <param name="Keys">Нормализованные заголовки, которые считаются этой же колонкой.</param>
public sealed record CreatedColumn(string Name, string Title, IReadOnlyList<string> Keys);
