using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;
using WarehousePlanAutomation.Core.Text;
using WarehousePlanAutomation.Tests.TestData;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class CellErrorReadingTests
{
    [Theory]
    [InlineData(CellError.NotAvailable)]
    [InlineData(CellError.DivideByZero)]
    [InlineData(CellError.Value)]
    [InlineData(CellError.Reference)]
    public void ОшибкаФормулыНеЧисло(int code)
    {
        Assert.Null(TextUtils.CellToDouble(code));
        Assert.Null(TextUtils.CellToDouble((double)code));
        Assert.Null(SheetGrid.FromRows(1, 1, new[] { new object?[] { code } }).Number(1, 1));
    }

    [Fact]
    public void ОбычныеЧислаЧитаютсяКакРаньше()
    {
        Assert.Equal(19021d, TextUtils.CellToDouble(19021d));
        Assert.Equal(-5d, TextUtils.CellToDouble(-5d));
        Assert.Equal(1.5d, TextUtils.CellToDouble("1,5"));
        Assert.Null(TextUtils.CellToDouble(string.Empty));
    }

    [Fact]
    public void ОшибкаВКолонкеНомераЗагрузкиНеСтановитсяНомером()
    {
        var source = PlanFixture.BuildGrid();
        var column = PlanSheetReader.Read(source).Headers[SheetSchema.Plan.LoadNumber];
        var urgentRow = PlanSheetReader.Read(source).Section(PlanSectionKind.AllGroups)!.DataRows[2].ExcelRow;

        var values = new object?[source.RowCount, source.ColumnCount];
        for (var r = 0; r < source.RowCount; r++)
        {
            for (var c = 0; c < source.ColumnCount; c++)
            {
                values[r, c] = source.Value(source.FirstRow + r, source.FirstColumn + c);
            }
        }

        values[urgentRow - source.FirstRow, column - source.FirstColumn] = CellError.NotAvailable;
        var grid = new SheetGrid(source.FirstRow, source.FirstColumn, values, null);

        var urgent = PlanSheetReader.Read(grid).Section(PlanSectionKind.AllGroups)!.DataRows[2];

        Assert.Null(urgent.LoadNumber);
    }
}
