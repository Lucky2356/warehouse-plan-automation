using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using WarehousePlanAutomation.Core.Text;
using WarehousePlanAutomation.Tests.TestData;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class PlanPriorityTests
{
    private static PlanSortItem Item(
        int index,
        double? networkDate,
        SetMonoKind kind = SetMonoKind.Neutral,
        bool urgent = false,
        bool acceptance = false,
        bool autoHub = false) =>
        new(index, acceptance, autoHub, urgent, networkDate, kind);

    private static List<PlanSortItem> Sort(IEnumerable<PlanSortItem> items, bool honorSpecialRows = true)
    {
        var list = items.ToList();
        list.Sort(new PlanPriorityComparer(honorSpecialRows));
        return list;
    }

    [Fact]
    public void ДатаВажнееСетИМоно()
    {
        var sorted = Sort(new[]
        {
            Item(0, 46280d, SetMonoKind.Set),
            Item(1, 46270d, SetMonoKind.Mono),
        });

        Assert.Equal(1, sorted[0].OriginalIndex);
        Assert.Equal(0, sorted[1].OriginalIndex);
    }

    [Fact]
    public void ПриОдинаковойДатеСетВышеМоно()
    {
        var sorted = Sort(new[]
        {
            Item(0, 46280d, SetMonoKind.Mono),
            Item(1, 46280d, SetMonoKind.Set),
        });

        Assert.Equal(1, sorted[0].OriginalIndex);
        Assert.Equal(0, sorted[1].OriginalIndex);
    }

    [Fact]
    public void СрочныеВышеНесрочныхДажеСПозднейДатой()
    {
        var sorted = Sort(new[]
        {
            Item(0, 46270d),
            Item(1, 46300d, urgent: true),
        });

        Assert.Equal(1, sorted[0].OriginalIndex);
    }

    [Fact]
    public void ОсобыеСтрокиВсегдаПервые()
    {
        var sorted = Sort(new[]
        {
            Item(0, 46270d, urgent: true),
            Item(1, null, autoHub: true),
            Item(2, null, acceptance: true),
        });

        Assert.Equal(2, sorted[0].OriginalIndex);
        Assert.Equal(1, sorted[1].OriginalIndex);
        Assert.Equal(0, sorted[2].OriginalIndex);
    }

    [Fact]
    public void ПустаяДатаОстаётсяВНачалеКатегории()
    {
        // Строке без «Даты в сети» значение не придумывается: отправить её вниз
        // значило бы считать её дату бесконечно поздней.
        var sorted = Sort(new[]
        {
            Item(0, 46300d),
            Item(1, null),
        });

        Assert.Equal(1, sorted[0].OriginalIndex);
        Assert.Equal(0, sorted[1].OriginalIndex);
    }

    [Fact]
    public void СтрокиБезДатыСохраняютВзаимныйПорядок()
    {
        var sorted = Sort(new[]
        {
            Item(0, 46300d),
            Item(1, null),
            Item(2, null),
            Item(3, 46250d),
        });

        Assert.Equal(new[] { 1, 2, 3, 0 }, sorted.Select(i => i.OriginalIndex));
    }

    [Fact]
    public void ПриПолномРавенствеПорядокСохраняется()
    {
        var sorted = Sort(new[]
        {
            Item(0, 46300d),
            Item(1, 46300d),
            Item(2, 46300d),
        });

        Assert.Equal(new[] { 0, 1, 2 }, sorted.Select(i => i.OriginalIndex));
    }

    [Fact]
    public void ПеремещенияПриводятБлокКЦелевомуПорядку()
    {
        var section = BuildSection(new[]
        {
            ("Приемка на хранилище от 24.08", (double?)null),
            ("Автозаказы для ХАБов", null),
            ("1100-026 Обувь МОНО", 46280d),
            ("Срочная подтоварка", 46300d),
            ("1100-026 Обувь СЕТ1", 46280d),
        });

        var moves = PlanArrangementBuilder.BuildSectionMoves(section, honorSpecialRows: true);
        var order = Simulate(section.DataRows.Select(r => r.Supplies).ToList(), section.FirstDataRow, moves);

        Assert.Equal(
            new[]
            {
                "Приемка на хранилище от 24.08",
                "Автозаказы для ХАБов",
                "Срочная подтоварка",
                "1100-026 Обувь СЕТ1",
                "1100-026 Обувь МОНО",
            },
            order);
    }

    [Fact]
    public void БлокСПустойСтрокойВнутри_НеСортируетсяНоНумеруется()
    {
        // Перемещения адресуют строки по смещению от начала блока, поэтому разрыв
        // внутри блока сделал бы адреса неверными: такой блок безопаснее не трогать.
        var layout = PlanFixture.BuildLayoutWithGapInAllGroups();
        var allGroups = layout.Section(PlanSectionKind.AllGroups)!;

        Assert.False(PlanArrangementBuilder.IsContiguous(allGroups));

        var arrangement = PlanArrangementBuilder.Build(layout);

        Assert.Empty(arrangement.Moves);
        Assert.Equal(
            allGroups.DataRows.Select(r => r.ExcelRow),
            arrangement.Numbers.Where(n => allGroups.DataRows.Any(r => r.ExcelRow == n.ExcelRow)).Select(n => n.ExcelRow));
    }

    private static PlanSection BuildSection(IReadOnlyList<(string Supplies, double? NetworkDate)> rows)
    {
        var dataRows = new List<PlanRow>();
        for (var i = 0; i < rows.Count; i++)
        {
            var supplies = rows[i].Supplies;
            dataRows.Add(new PlanRow(3 + i, PlanSectionKind.AllGroups)
            {
                Supplies = supplies,
                NetworkDate = rows[i].NetworkDate,
                IsStorageAcceptanceRow = TextUtils.StartsWithKey(supplies, "приемка на хранилище"),
                IsAutoHubRow = TextUtils.StartsWithKey(supplies, "автозаказы для хабов"),
            });
        }

        return new PlanSection(PlanSectionKind.AllGroups, 2, "=SUM(J5:J7)", dataRows);
    }

    /// <summary>Повторяет поведение Excel: вырезанная строка вставляется перед целевой.</summary>
    private static List<string> Simulate(List<string> order, int firstRow, IReadOnlyList<RowMove> moves)
    {
        foreach (var move in moves)
        {
            var from = move.FromRow - firstRow;
            var to = move.ToRow - firstRow;
            var value = order[from];
            order.Insert(to, value);
            order.RemoveAt(from + 1);
        }

        return order;
    }
}
