using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class MarkupAssignmentTests
{
    private static readonly IReadOnlyList<MarkupRule> Rules = new[]
    {
        new MarkupRule("БИЖУТЕРИЯ", "БИЖУТЕРИЯ", 5.69d, 4.50d),
        new MarkupRule("ОБУВЬ", "ОБУВЬ зима", 3.10d, 2.45d),
        new MarkupRule("ОБУВЬ", "ОБУВЬ лето", 3.40d, 2.69d),
    };

    private sealed class Answers : IDecisionPrompt
    {
        private readonly Func<DecisionRequest, string?> _answer;

        public Answers(Func<DecisionRequest, string?> answer) => _answer = answer;

        public int Calls { get; private set; }

        public string? Choose(DecisionRequest request)
        {
            Calls++;
            return _answer(request);
        }
    }

    private static PriceRowValues Row(int index, string sector, string subgroup) => new(
        index,
        0,
        "18682802" + index,
        "431-329",
        10d,
        6.55d,
        65.5d,
        6.55d,
        new PriceListRow("18682802" + index, "1555" + index, "модель", sector, "ГРУППА",
            "наименование", "431-2802", "ЧЕРНЫЙ", "WINTER 26-27", "TU"),
        subgroup,
        null,
        false);

    [Fact]
    public void КаждаяПодгруппаПолучаетСвоюНаценку()
    {
        var prompt = new Answers(request =>
            request.Options.First(o => o.Contains(request.Question.Contains("САПОГИ") ? "зима" : "лето")));

        var plan = MarkupAssignment.Apply(
            new[]
            {
                Row(0, "ОБУВЬ", "САПОГИ/"),
                Row(1, "ОБУВЬ", "БОСОНОЖКИ/"),
                Row(2, "ОБУВЬ", "САПОГИ/"),
            },
            new MarkupResolver(Rules, prompt));

        Assert.Equal(2, prompt.Calls);
        Assert.Equal(new double?[] { 3.10d, 3.40d, 3.10d }, plan.Rows.Select(row => row.Markup!.Planned));
        Assert.Empty(plan.Warnings);
    }

    [Fact]
    public void ОдинВариантВРегламенте_НичегоНеСпрашивает()
    {
        var prompt = new Answers(_ => null);

        var plan = MarkupAssignment.Apply(
            new[] { Row(0, "БИЖУТЕРИЯ", "СЕРЬГИ/"), Row(1, "БИЖУТЕРИЯ", "КОЛЬЦА/") },
            new MarkupResolver(Rules, prompt));

        Assert.Equal(0, prompt.Calls);
        Assert.All(plan.Rows, row => Assert.Equal(5.69d, row.Markup!.Planned));
    }

    [Fact]
    public void НеизвестныйСектор_ОставляетНаценкиПустымиИПишетЗамечаниеОдинРаз()
    {
        var plan = MarkupAssignment.Apply(
            new[] { Row(0, "ЗОНТЫ", "ЗОНТ/"), Row(1, "ЗОНТЫ", "ЗОНТ/") },
            new MarkupResolver(Rules));

        Assert.All(plan.Rows, row => Assert.Null(row.Markup));
        var warning = Assert.Single(plan.Warnings);
        Assert.Contains("ЗОНТЫ", warning.Message);
    }

    [Fact]
    public void НеизвестныйСектор_НаценкуВписаннуюРуками_НеСтирает()
    {
        // Пересчёт запускают не один раз: то, что человек вписал после первого запуска,
        // второй запуск затирать не должен.
        var byHand = new MarkupRule("ЗОНТЫ", string.Empty, 4.20d, 3.30d);

        var plan = MarkupAssignment.Apply(
            new[] { Row(0, "ЗОНТЫ", "ЗОНТ/") with { Markup = byHand } },
            new MarkupResolver(Rules));

        Assert.Same(byHand, Assert.Single(plan.Rows).Markup);
        Assert.Single(plan.Warnings);
    }

    [Fact]
    public void ВыборНеСделан_НаценкуВписаннуюРуками_НеСтирает()
    {
        var byHand = new MarkupRule("ОБУВЬ", string.Empty, 3.25d, 2.58d);
        var prompt = new Answers(_ => null);

        var plan = MarkupAssignment.Apply(
            new[] { Row(0, "ОБУВЬ", "САПОГИ/") with { Markup = byHand } },
            new MarkupResolver(Rules, prompt));

        Assert.Equal(1, prompt.Calls);
        Assert.Same(byHand, Assert.Single(plan.Rows).Markup);
    }
}
