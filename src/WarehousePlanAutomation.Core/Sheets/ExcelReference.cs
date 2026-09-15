using System.Globalization;
using System.Text.RegularExpressions;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>Адрес одной ячейки в абсолютных координатах Excel.</summary>
public readonly record struct CellRef(int Row, int Column)
{
    public override string ToString() =>
        ExcelColumn.ToLetters(Column) + Row.ToString(CultureInfo.InvariantCulture);
}

/// <summary>Адрес прямоугольного диапазона.</summary>
public readonly record struct RangeRef(int FirstRow, int FirstColumn, int LastRow, int LastColumn)
{
    public int RowCount => LastRow - FirstRow + 1;

    public int ColumnCount => LastColumn - FirstColumn + 1;

    public override string ToString() =>
        new CellRef(FirstRow, FirstColumn) + ":" + new CellRef(LastRow, LastColumn);
}

/// <summary>
/// Разбор ссылок из текста формулы.
///
/// Нужен там, где разметку листа задаёт не строка заголовков, а сами формулы: на листе
/// «Распред» блок сезонности и ячейка сектора нигде не подписаны, но обе стоят в формулах
/// таблицы РТТ. Читать их оттуда точнее, чем угадывать по расположению.
/// </summary>
public static class ExcelReference
{
    private const string Cell = @"\$?(?<col>[A-Za-z]{1,3})\$?(?<row>\d{1,7})";

    private static readonly Regex CellPattern = new(
        "^" + Cell + "$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex RangePattern = new(
        @"\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>\d{1,7})\s*:\s*\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>\d{1,7})",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Ссылка на ячейку, приклеенная к TEXT(...) через «&». Так в таблице РТТ собирается
    /// ключ поиска: TEXT($A28;"000")&$H$26 - и вторая половина ключа это ячейка сектора.
    /// </summary>
    private static readonly Regex ConcatenatedCellPattern = new(
        @"TEXT\s*\([^()]*\)\s*&\s*" + Cell,
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static bool TryParseCell(string? text, out CellRef cell)
    {
        cell = default;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var match = CellPattern.Match(text.Trim());
        if (!match.Success)
        {
            return false;
        }

        cell = new CellRef(
            int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture),
            ExcelColumn.FromLetters(match.Groups["col"].Value));
        return true;
    }

    /// <summary>Ячейка, приписанная к ключу поиска через «&». null, если такой нет.</summary>
    public static CellRef? FindConcatenatedCell(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return null;
        }

        var match = ConcatenatedCellPattern.Match(formula);
        return match.Success
            ? new CellRef(
                int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture),
                ExcelColumn.FromLetters(match.Groups["col"].Value))
            : null;
    }

    /// <summary>
    /// Первый диапазон формулы, который лежит на этом же листе. Ссылки на другие листы
    /// («Цены!$D:$H») пропускаются: у них разметка своя.
    /// </summary>
    public static RangeRef? FindLocalRange(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return null;
        }

        foreach (Match match in RangePattern.Matches(formula))
        {
            if (IsQualified(formula, match.Index))
            {
                continue;
            }

            var first = new CellRef(
                int.Parse(match.Groups["r1"].Value, CultureInfo.InvariantCulture),
                ExcelColumn.FromLetters(match.Groups["c1"].Value));
            var last = new CellRef(
                int.Parse(match.Groups["r2"].Value, CultureInfo.InvariantCulture),
                ExcelColumn.FromLetters(match.Groups["c2"].Value));

            return new RangeRef(
                Math.Min(first.Row, last.Row),
                Math.Min(first.Column, last.Column),
                Math.Max(first.Row, last.Row),
                Math.Max(first.Column, last.Column));
        }

        return null;
    }

    /// <summary>Перед ссылкой стоит «!» - значит, она относится к другому листу.</summary>
    private static bool IsQualified(string formula, int index)
    {
        for (var i = index - 1; i >= 0; i--)
        {
            var ch = formula[i];
            if (ch == '!')
            {
                return true;
            }

            if (!char.IsLetterOrDigit(ch) && ch != '$' && ch != '\'' && ch != ' ' && ch != '_' && ch != '.')
            {
                return false;
            }
        }

        return false;
    }
}
