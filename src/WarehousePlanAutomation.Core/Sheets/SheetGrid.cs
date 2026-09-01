using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Прочитанный в память прямоугольник листа: значения (Value2) и, при необходимости, формулы.
/// Индексация ведётся абсолютными координатами Excel (строка и колонка начинаются с 1).
/// </summary>
public sealed class SheetGrid
{
    private readonly object?[,] _values;
    private readonly string?[,]? _formulas;

    public SheetGrid(int firstRow, int firstColumn, object?[,] values, string?[,]? formulas = null)
    {
        if (firstRow < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(firstRow));
        }

        if (firstColumn < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(firstColumn));
        }

        FirstRow = firstRow;
        FirstColumn = firstColumn;
        _values = values;
        _formulas = formulas;
        RowCount = values.GetLength(0);
        ColumnCount = values.GetLength(1);

        if (formulas is not null &&
            (formulas.GetLength(0) != RowCount || formulas.GetLength(1) != ColumnCount))
        {
            throw new ArgumentException("Размеры массивов значений и формул не совпадают.", nameof(formulas));
        }
    }

    public int FirstRow { get; }

    public int FirstColumn { get; }

    public int RowCount { get; }

    public int ColumnCount { get; }

    public int LastRow => FirstRow + RowCount - 1;

    public int LastColumn => FirstColumn + ColumnCount - 1;

    public bool Contains(int row, int column) =>
        row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;

    public object? Value(int row, int column) =>
        Contains(row, column) ? _values[row - FirstRow, column - FirstColumn] : null;

    public string Text(int row, int column) => TextUtils.CellToString(Value(row, column));

    public string NormalizedText(int row, int column) => TextUtils.Normalize(Text(row, column));

    public double? Number(int row, int column) => TextUtils.CellToDouble(Value(row, column));

    public string? Formula(int row, int column) =>
        _formulas is not null && Contains(row, column) ? _formulas[row - FirstRow, column - FirstColumn] : null;

    public bool HasFormula(int row, int column)
    {
        var formula = Formula(row, column);
        return formula is not null && formula.StartsWith("=", StringComparison.Ordinal);
    }

    /// <summary>Удобный конструктор для тестов и для данных, собранных построчно.</summary>
    public static SheetGrid FromRows(int firstRow, int firstColumn, IReadOnlyList<object?[]> rows)
    {
        var columnCount = rows.Count == 0 ? 0 : rows.Max(r => r.Length);
        var values = new object?[rows.Count, columnCount];
        for (var r = 0; r < rows.Count; r++)
        {
            for (var c = 0; c < rows[r].Length; c++)
            {
                values[r, c] = rows[r][c];
            }
        }

        return new SheetGrid(firstRow, firstColumn, values);
    }
}
