namespace WarehousePlanAutomation.Core.Text;

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

    /// <summary>
    /// Заказ под приемку на хранилище. «ё» ключ уже привёл к «е», так что «приёмка»
    /// находится так же.
    /// </summary>
    public const string StorageAcceptanceMarker = "приемк";

    public const string LoadedComment = "Заказы загружены.";
    public const string PlaceholderCommentMarker = "заказы будут загружены";
    public const string InAssemblyStatus = "в сборке";
    public const string CrossDockProcessing = "перекр";
    public const string JournalStartedStatus = "запущен";

    /// <summary>Документ «ЗП» в комментарии выгрузки, например «ЗП373-1739».</summary>
    public const string ZpDocumentMarker = "зп";

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
    /// обычный магазин «Магнитогорск-М66». По той же причине как отдельные слова
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

    /// <summary>
    /// Строка документа «ЗП»: такие строки удаляются с листа выгрузки и в план не переносятся.
    /// Признак ищется как отдельный, а не как подстрока: «ЗП373-1739» находится,
    /// а «зп» внутри другого слова - нет.
    /// </summary>
    public static bool IsZpRow(string? comment) => TextUtils.ContainsToken(comment, ZpDocumentMarker);

    public static bool IsReturn(string? text) => TextUtils.ContainsKey(text, ReturnsMarker);

    public static bool IsStorageAcceptance(string? text) => TextUtils.ContainsKey(text, StorageAcceptanceMarker);

    /// <summary>«Срочная подтоварка 28.08_Хранение»: у такой строки «Дата в сети» на 4 дня раньше.</summary>
    public static bool IsUrgentRestock(string? text) =>
        TextUtils.ContainsKey(text, UrgentMarker) && TextUtils.ContainsKey(text, RestockMarker);

    public const string UrgentMarker = "срочн";

    public const string RestockMarker = "подтовар";

    /// <summary>На сколько дней раньше срочная подтоварка должна быть в сети.</summary>
    public const int UrgentRestockDays = 4;

    /// <summary>
    /// Формула «Дата в сети» для новой строки. В плане это «=F5+$O$2», а у срочной подтоварки -
    /// «=F5+$O$2-4». Формула берётся у соседней строки, у которой вычитание может быть,
    /// а может и не быть, поэтому оно сначала убирается, а потом ставится, если нужно.
    /// </summary>
    public static string NetworkDateFormula(string donorFormula, bool urgentRestock)
    {
        var trimmed = System.Text.RegularExpressions.Regex.Replace(
            donorFormula.TrimEnd(), @"\s*-\s*" + UrgentRestockDays + "$", string.Empty);
        return urgentRestock ? trimmed + "-" + UrgentRestockDays : trimmed;
    }

    public static bool IsPlaceholderComment(string? text) =>
        TextUtils.ContainsKey(text, PlaceholderCommentMarker);

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
