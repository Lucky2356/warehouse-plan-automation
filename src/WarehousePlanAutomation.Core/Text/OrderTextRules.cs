namespace WarehousePlanAutomation.Core.Text;

/// <summary>Признак «СЕТ» / «МОНО» в колонке «Поставки».</summary>
public enum SetMonoKind
{
    Set = 0,
    Neutral = 1,
    Mono = 2,
}

/// <summary>Способ сопоставления названия подразделения с признаком.</summary>
public enum MarkerMatch
{
    /// <summary>Вхождение подстроки в любом месте названия.</summary>
    Contains,

    /// <summary>Вхождение как отдельного слова.</summary>
    Word,
}

/// <summary>Признак подразделения: искомый текст и способ сопоставления.</summary>
public sealed record DivisionMarker(string Text, MarkerMatch Match)
{
    public bool Matches(string? division) => Match == MarkerMatch.Word
        ? TextUtils.ContainsWord(division, Text)
        : TextUtils.ContainsKey(division, Text);
}

/// <summary>Текстовые признаки, зафиксированные техническим заданием.</summary>
public static class OrderTextRules
{
    public const string ReturnsMarker = "возвр";
    public const string UrgencyMarker = "сроч";
    public const string SetMarker = "сет";
    public const string MonoMarker = "моно";
    public const string LoadedComment = "Заказы загружены.";
    public const string PlaceholderCommentMarker = "заказы будут загружены";
    public const string InAssemblyStatus = "в сборке";
    public const string CrossDockProcessing = "перекр";
    public const string JournalStartedStatus = "запущен";

    /// <summary>
    /// Строки, которые удаляются без переноса куда-либо.
    /// «виртуал» перекрывается признаком «вирт», но указан отдельно,
    /// чтобы список дословно соответствовал согласованному перечню.
    /// </summary>
    public static readonly IReadOnlyList<string> ServiceMarkers = new[]
    {
        "автозаказ",
        "вирт",
        "виртуал",
        "фото",
        "образцы",
        "ремонт",
        "списать",
        "маркировка",
    };

    /// <summary>Подразделение «Опт».</summary>
    public static readonly IReadOnlyList<DivisionMarker> WholesaleMarkers = new[]
    {
        new DivisionMarker("опт", MarkerMatch.Contains),
    };

    /// <summary>Подразделение интернет-магазина.</summary>
    public static readonly IReadOnlyList<DivisionMarker> InternetShopMarkers = new[]
    {
        new DivisionMarker("777", MarkerMatch.Contains),
    };

    /// <summary>
    /// Маркетплейсы. Кроме написаний из ТЗ учтены латинские варианты, которыми
    /// выгрузка называет те же подразделения (Ozon, Lamoda, Wildberries, Sber).
    ///
    /// «Магнит» сопоставляется как отдельное слово: иначе под правило попал бы
    /// обычный магазин «Магнитогорск-М96». По той же причине как отдельные слова
    /// сопоставляются короткие «ВБ» и «WB».
    /// </summary>
    public static readonly IReadOnlyList<DivisionMarker> MarketplaceMarkers = new[]
    {
        new DivisionMarker("озон", MarkerMatch.Contains),
        new DivisionMarker("ozon", MarkerMatch.Contains),
        new DivisionMarker("вб", MarkerMatch.Word),
        new DivisionMarker("wb", MarkerMatch.Word),
        new DivisionMarker("wildberries", MarkerMatch.Contains),
        new DivisionMarker("ламода", MarkerMatch.Contains),
        new DivisionMarker("lamoda", MarkerMatch.Contains),
        new DivisionMarker("сбер", MarkerMatch.Contains),
        new DivisionMarker("sber", MarkerMatch.Contains),
        new DivisionMarker("екатеринбург яблоко", MarkerMatch.Contains),
        new DivisionMarker("магнит", MarkerMatch.Word),
        new DivisionMarker("magnit", MarkerMatch.Word),
    };

    public static bool IsWholesale(string? division) => Matches(division, WholesaleMarkers);

    public static bool IsInternetShop(string? division) => Matches(division, InternetShopMarkers);

    public static bool IsMarketplace(string? division) => Matches(division, MarketplaceMarkers);

    public static bool IsServiceRow(string? comment) => TextUtils.ContainsAnyKey(comment, ServiceMarkers);

    public static bool IsReturn(string? text) => TextUtils.ContainsKey(text, ReturnsMarker);

    public static bool IsUrgent(string? text) => TextUtils.ContainsKey(text, UrgencyMarker);

    public static bool IsPlaceholderComment(string? text) =>
        TextUtils.ContainsKey(text, PlaceholderCommentMarker);

    public static SetMonoKind DetectSetMono(string? supplies)
    {
        if (TextUtils.ContainsKey(supplies, SetMarker))
        {
            return SetMonoKind.Set;
        }

        return TextUtils.ContainsKey(supplies, MonoMarker) ? SetMonoKind.Mono : SetMonoKind.Neutral;
    }

    private static bool Matches(string? division, IReadOnlyList<DivisionMarker> markers)
    {
        foreach (var marker in markers)
        {
            if (marker.Matches(division))
            {
                return true;
            }
        }

        return false;
    }
}
