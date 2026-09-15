using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Processing;

/// <summary>Строка листа «Аналоги»: АЦ и его аналог - «последний АЦ, который приходит в сеть».</summary>
public sealed record AnaloguePair(string Vsp, string Analogue);

/// <summary>
/// «Аналог гугл» по листу «Аналоги».
///
/// Вручную это два ВПР на листе «Аналоги» против колонки «Всп» листа «Цены», фильтры,
/// зелёная заливка и перенос строк в свободные колонки J и K. Итог процедуры такой:
/// аналог есть у того АЦ поставки, который стоит в листе справа (в колонке «Аналог»),
/// а АЦ слева в этой строке в поставку не входит. Аналогом тогда считается АЦ слева.
/// Строки, где АЦ поставки стоит слева, при ручной разметке красятся зелёным и в J:K не попадают.
///
/// Всем остальным АЦ поставки пишется «нет» - так же, как «#н/д» заменяется на «нет» вручную.
/// </summary>
public static class AnalogueLookup
{
    public const string None = "нет";

    /// <summary>
    /// Аналог для каждого АЦ поставки. Ключ - нормализованный АЦ; если строк с одним АЦ
    /// справа несколько, берётся первая - как это сделал бы ВПР.
    /// </summary>
    public static IReadOnlyDictionary<string, string> Build(
        IEnumerable<AnaloguePair> pairs,
        IEnumerable<string> supplyVsp)
    {
        var supply = supplyVsp
            .Select(TextUtils.NormalizeKey)
            .Where(key => key.Length > 0)
            .ToHashSet(StringComparer.Ordinal);

        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in pairs)
        {
            var left = TextUtils.NormalizeKey(pair.Vsp);
            var right = TextUtils.NormalizeKey(pair.Analogue);
            if (left.Length == 0 || right.Length == 0)
            {
                continue;
            }

            if (supply.Contains(right) && !supply.Contains(left))
            {
                result.TryAdd(right, TextUtils.Normalize(pair.Vsp));
            }
        }

        return result;
    }

    /// <summary>Значение «Аналог гугл» для АЦ: название аналога либо «нет».</summary>
    public static string ValueFor(IReadOnlyDictionary<string, string> analogues, string? vsp) =>
        analogues.TryGetValue(TextUtils.NormalizeKey(vsp), out var analogue) ? analogue : None;
}
