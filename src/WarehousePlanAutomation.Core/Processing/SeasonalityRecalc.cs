using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>
/// Пересчёт долей сезонности для секторов, которым его делают руками по шаблону.
///
/// В блок «Сезонность по НК» всегда идут три месяца, посчитанные с листа «КС». Дальше
/// для трёх секторов доли пересчитываются: доля месяца делится на сумму долей за месяцы,
/// которые участвуют в расчёте - 4 % и 6 % превращаются в 40 % и 60 %. У бижутерии таких
/// месяцев два, у колготок с носками и украшений для волос - три. Третий месяц бижутерии
/// остаётся таким, как посчитался с «КС»: он справочный.
///
/// Для остальных секторов доли остаются как есть.
/// </summary>
public static class SeasonalityRecalc
{
    /// <summary>Секторы написаны так же, как в колонке «Сектор» «Сводного прайса».</summary>
    public const string JewellerySector = "БИЖУТЕРИЯ";

    public const string SocksSector = "КОЛГОТКИ,НОСКИ";

    public const string HairSector = "УКРАШЕНИЯ ДЛЯ ВОЛОС";

    /// <summary>Сколько месяцев пересчитывается для сектора. 0 - сектор не пересчитывается.</summary>
    public static int MonthsFor(string? sector)
    {
        var key = Key(sector);
        if (key.Length == 0)
        {
            return 0;
        }

        if (key == Key(JewellerySector))
        {
            return 2;
        }

        return key == Key(SocksSector) || key == Key(HairSector) ? 3 : 0;
    }

    /// <summary>
    /// Делит доли первых <paramref name="months"/> месяцев на их сумму. Остальные месяцы
    /// не трогаются. Если сумма нулевая или пустая, доли остаются как были: делить не на что.
    /// </summary>
    public static IReadOnlyList<double?> Apply(IReadOnlyList<double?> shares, int months)
    {
        if (months <= 0 || shares.Count == 0)
        {
            return shares;
        }

        var count = Math.Min(months, shares.Count);
        var total = 0d;
        for (var i = 0; i < count; i++)
        {
            total += shares[i] ?? 0d;
        }

        if (total <= 0d)
        {
            return shares;
        }

        var result = shares.ToList();
        for (var i = 0; i < count; i++)
        {
            result[i] = (shares[i] ?? 0d) / total;
        }

        return result;
    }

    /// <summary>
    /// Название сектора без пробелов: в выгрузках встречается и «КОЛГОТКИ,НОСКИ»,
    /// и «КОЛГОТКИ, НОСКИ».
    /// </summary>
    private static string Key(string? sector) =>
        TextUtils.NormalizeKey(sector).Replace(" ", string.Empty, StringComparison.Ordinal);
}
