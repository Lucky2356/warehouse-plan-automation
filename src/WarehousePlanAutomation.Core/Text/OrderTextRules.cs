namespace WarehousePlanAutomation.Core.Text;

/// <summary>Признак «СЕТ» / «МОНО» в колонке «Поставки».</summary>
public enum SetMonoKind
{
    Set = 0,
    Neutral = 1,
    Mono = 2,
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

    /// <summary>Строки, которые удаляются без переноса куда-либо.</summary>
    public static readonly IReadOnlyList<string> ServiceMarkers = new[] { "автозаказ", "вирт", "фото" };

    /// <summary>Подразделение «Опт».</summary>
    public static readonly IReadOnlyList<string> WholesaleMarkers = new[] { "опт" };

    /// <summary>Подразделение интернет-магазина.</summary>
    public static readonly IReadOnlyList<string> InternetShopMarkers = new[] { "777" };

    /// <summary>
    /// Маркетплейсы. Кроме написаний из ТЗ учтены латинские варианты, которыми
    /// выгрузка называет те же подразделения (Ozon, Lamoda, Wildberries, Sber).
    /// </summary>
    public static readonly IReadOnlyList<string> MarketplaceMarkers = new[]
    {
        "озон", "ozon",
        "вб", "wb", "wildberries",
        "ламода", "lamoda",
        "сбер", "sber",
        "екатеринбург яблоко",
    };

    public static bool IsWholesale(string? division) => TextUtils.ContainsAnyKey(division, WholesaleMarkers);

    public static bool IsInternetShop(string? division) => TextUtils.ContainsAnyKey(division, InternetShopMarkers);

    public static bool IsMarketplace(string? division) => TextUtils.ContainsAnyKey(division, MarketplaceMarkers);

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
}
