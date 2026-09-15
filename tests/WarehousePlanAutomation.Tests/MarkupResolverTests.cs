using WarehousePlanAutomation.Core.Abstractions;
using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Processing;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class MarkupResolverTests
{
    private static readonly IReadOnlyList<MarkupRule> Rules = new[]
    {
        new MarkupRule("БИЖУТЕРИЯ", "БИЖУТЕРИЯ", 5.69d, 4.50d),
        new MarkupRule("ОБУВЬ", "ОБУВЬ зима", 3.10d, 2.45d),
        new MarkupRule("ОБУВЬ", "ОБУВЬ лето", 3.40d, 2.69d),
        new MarkupRule("МЕЛКИЕ АКСЕССУАРЫ", "ПРОЧЕЕ", 5.24d, 4.14d),
    };

    private sealed class StubPrompt : IDecisionPrompt
    {
        private readonly Func<DecisionRequest, string?> _answer;

        public StubPrompt(Func<DecisionRequest, string?> answer) => _answer = answer;

        public int Calls { get; private set; }

        public string? Choose(DecisionRequest request)
        {
            Calls++;
            return _answer(request);
        }
    }

    [Fact]
    public void ЕдинственныйСектор_БеретсяБезВопросов()
    {
        var prompt = new StubPrompt(_ => null);
        var choice = new MarkupResolver(Rules, prompt).Resolve("БИЖУТЕРИЯ");

        Assert.Equal(0, prompt.Calls);
        Assert.Equal(5.69d, choice.Rule!.Planned);
        Assert.Equal(4.50d, choice.Rule.Minimum);
        Assert.Null(choice.Warning);
    }

    [Fact]
    public void РегистрИЛишниеПробелыНеМешают()
    {
        var choice = new MarkupResolver(Rules).Resolve("  бижутерия ");

        Assert.Equal(5.69d, choice.Rule!.Planned);
    }

    [Fact]
    public void НесколькоСтрокСектора_СпрашиваетЧеловека()
    {
        var prompt = new StubPrompt(request => request.Options.First(o => o.Contains("лето")));
        var choice = new MarkupResolver(Rules, prompt).Resolve("ОБУВЬ");

        Assert.Equal(1, prompt.Calls);
        Assert.Equal(3.40d, choice.Rule!.Planned);
        Assert.Null(choice.Warning);
    }

    [Fact]
    public void ВопросЗадаётсяОдинРазНаПодгруппу()
    {
        var prompt = new StubPrompt(request => request.Options[0]);
        var resolver = new MarkupResolver(Rules, prompt);

        resolver.Resolve("ОБУВЬ", "ТАПОЧКИ/");
        resolver.Resolve("ОБУВЬ", "ТАПОЧКИ/");
        resolver.Resolve("обувь", " тапочки/ ");

        Assert.Equal(1, prompt.Calls);
    }

    [Fact]
    public void РазныеПодгруппы_СпрашиваютсяПоОтдельности()
    {
        // В одной поставке подгрупп может быть несколько, и наценка у каждой своя.
        var asked = new List<string>();
        var prompt = new StubPrompt(request =>
        {
            asked.Add(request.Question);
            return request.Options.First(o => o.Contains(asked.Count == 1 ? "зима" : "лето"));
        });
        var resolver = new MarkupResolver(Rules, prompt);

        var boots = resolver.Resolve("ОБУВЬ", "САПОГИ/");
        var sandals = resolver.Resolve("ОБУВЬ", "БОСОНОЖКИ/");

        Assert.Equal(2, prompt.Calls);
        Assert.Equal(3.10d, boots.Rule!.Planned);
        Assert.Equal(3.40d, sandals.Rule!.Planned);
        Assert.Contains("САПОГИ/", asked[0]);
        Assert.Contains("БОСОНОЖКИ/", asked[1]);
    }

    [Fact]
    public void ОтказОтВыбора_ОставляетНаценкиПустымиИПишетЗамечание()
    {
        var choice = new MarkupResolver(Rules, new StubPrompt(_ => null)).Resolve("ОБУВЬ");

        Assert.Null(choice.Rule);
        Assert.Contains("несколько строк", choice.Warning);
        Assert.Contains("ОБУВЬ зима", choice.Warning);
    }

    [Fact]
    public void БезСобеседника_НеУгадывает()
    {
        // Программа не выбирает строку сама: в консольном прогоне без интерфейса
        // наценки остаются пустыми, а не берутся «первые попавшиеся».
        var choice = new MarkupResolver(Rules).Resolve("ОБУВЬ");

        Assert.Null(choice.Rule);
        Assert.NotNull(choice.Warning);
    }

    [Fact]
    public void НеизвестныйСектор_ДаётЗамечание()
    {
        var choice = new MarkupResolver(Rules).Resolve("ЗОНТЫ");

        Assert.Null(choice.Rule);
        Assert.Contains("ЗОНТЫ", choice.Warning);
    }

    /// <summary>Строки «Регламента наценок» из примера распреда.</summary>
    private static readonly IReadOnlyList<MarkupRule> Regulation = new[]
    {
        new MarkupRule("МЕЛКИЕ АКСЕССУАРЫ", "ПРОЧЕЕ", 5.24d, 4.14d),
        new MarkupRule("МЕЛКИЕ АКСЕССУАРЫ", "МЕЛ АКСЕСС под упак", 5.80d, 4.58d),
        new MarkupRule("ПЕРЧАТКИ", "РУКАВИЦЫ", 5.18d, 4.09d),
        new MarkupRule("ПЕРЧАТКИ", "ПЕРЧАТКИ", 5.59d, 4.42d),
        new MarkupRule("СУМКИ", "СУМКИ рюкзак", 3.67d, 2.90d),
        new MarkupRule("СУМКИ", "СУМКИ прочее", 3.63d, 2.87d),
        new MarkupRule("СУМКИ", "СУМКИ мешок", 2.57d, 2.03d),
        new MarkupRule("СУМКИ", "СУМКИ кошелек", 3.97d, 3.14d),
        new MarkupRule("ОБУВЬ", "ОБУВЬ тапочки", 4.15d, 3.28d),
        new MarkupRule("ОБУВЬ", "ОБУВЬ зима", 3.27d, 2.59d),
        new MarkupRule("ОБУВЬ", "ОБУВЬ лето", 2.31d, 1.83d),
        new MarkupRule("ШАРФЫ И ПЛАТКИ", "ШАРФЫ И ПЛАТКИлето", 7.00d, 5.54d),
        new MarkupRule("ШАРФЫ И ПЛАТКИ", "ШАРФЫ И ПЛАТКИзима", 5.22d, 4.13d),
    };

    [Fact]
    public void ПодгруппаНеСовпалаСоСтрокой_БерётсяПрочее()
    {
        // Пример из распреда: маска для сна - не «под упак», наценка стоит по «ПРОЧЕЕ».
        var prompt = new StubPrompt(_ => null);
        var choice = new MarkupResolver(Regulation, prompt).Resolve("МЕЛКИЕ АКСЕССУАРЫ", "МАСКА ДЛЯ СНА/");

        Assert.Equal(0, prompt.Calls);
        Assert.Equal(5.24d, choice.Rule!.Planned);
        Assert.Null(choice.Warning);
    }

    [Theory]
    [InlineData("СУМКИ", "РЮКЗАК ГОРОДСКОЙ/", 3.67d)]
    [InlineData("СУМКИ", "КОШЕЛЬКИ/", 3.97d)]
    [InlineData("СУМКИ", "СУМКА-МЕШОК/", 2.57d)]
    [InlineData("СУМКИ", "ШОППЕР/", 3.63d)]
    [InlineData("ПЕРЧАТКИ", "МАРИНА/рукавицы", 5.18d)]
    [InlineData("ПЕРЧАТКИ", "Леди/перчатки", 5.59d)]
    [InlineData("ОБУВЬ", "ТАПОЧКИ/", 4.15d)]
    public void СтрокаВыбираетсяПоПодгруппе(string sector, string subgroup, double planned)
    {
        var prompt = new StubPrompt(_ => null);
        var choice = new MarkupResolver(Regulation, prompt).Resolve(sector, subgroup);

        Assert.Equal(0, prompt.Calls);
        Assert.Equal(planned, choice.Rule!.Planned);
    }

    [Fact]
    public void НиПодгруппаНиПрочееНеПодошли_СпрашиваетСредиВсехСтрок()
    {
        DecisionRequest? asked = null;
        var prompt = new StubPrompt(request =>
        {
            asked = request;
            return request.Options.First(o => o.Contains("зима"));
        });

        var choice = new MarkupResolver(Regulation, prompt).Resolve("ОБУВЬ", "САПОГИ/");

        Assert.Equal(1, prompt.Calls);
        Assert.Equal(3, asked!.Options.Count);
        Assert.Equal(3.27d, choice.Rule!.Planned);
    }

    [Fact]
    public void СлитноеНаписаниеСектораПлюс_РазбираетсяНаСлова()
    {
        Assert.Equal(new[] { "шарфы", "и", "платки", "лето" }, MarkupSubgroupMatch.Words("ШАРФЫ И ПЛАТКИлето"));
        Assert.Equal(new[] { "леди", "перчатки" }, MarkupSubgroupMatch.Words("Леди\\перчатки"));
    }

    [Fact]
    public void ПустойСектор_ДаётЗамечание()
    {
        var choice = new MarkupResolver(Rules).Resolve("   ");

        Assert.Null(choice.Rule);
        Assert.Contains("не заполнен сектор", choice.Warning);
    }
}
