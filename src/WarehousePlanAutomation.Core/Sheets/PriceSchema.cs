namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Названия листов и колонок книги распреда. Как и в плане склада, буквы колонок
/// и номера строк нигде не зашиты: всё ищется по заголовкам.
///
/// Названия листов задаются началом названия - лист может называться «Прайс по
/// подразделениям» или «Прайс по подраздел.», и оба варианта правильные.
/// </summary>
public static class PriceSchema
{
    public const string InvoiceSheet = "Invoice";
    public const string ForPricesSheet = "для цен";
    public const string PricesSheet = "Цены";
    public const string SummaryPriceSheet = "Сводный прайс";
    public const string DivisionPriceSheet = "Прайс по подраздел";
    public const string MarkupSheet = "Регламент наценок";
    public const string LinkSheet = "link";
    public const string StockSourceSheet = "Остатки Н";
    public const string StockSheet = "остатки";
    public const string DistributionSheet = "Распред";
    public const string LoaderSheet = "Загрузочник";
    public const string SeasonalitySheet = "КС";
    public const string AnaloguesSheet = "Аналоги";

    /// <summary>
    /// Лист «Остатки Н» - выгрузка остатков по всем секторам. Программа его не меняет:
    /// «АЦР» собирается в памяти, лишние колонки в памяти же отбрасываются.
    /// </summary>
    public static class StockSource
    {
        public const string Acr = "АЦР";
        public const string Sector = "Сектор";
        public const string Article = "Артикул";
        public const string Color = "Цвет";
        public const string Size = "Размер";

        /// <summary>
        /// Колонки, которые в «остатки» не переносятся. «АЦ» заменяется на «АЦР»,
        /// остальные к распределению отношения не имеют.
        /// </summary>
        public static readonly IReadOnlyList<string> Dropped = new[]
        {
            "ац",
            "резерв",
            "без маркировки",
            "остаток маркетплейс",
            "остаток нет маркировки",
        };

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Sector, new[] { "сектор" }, exactOnly: true),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),
        };

        /// <summary>
        /// Строки магазинов на листе «остатки» называются «001 тз», «001 в_пути», «001 прод».
        /// В расчёт остатка идут первые две: проданное остатком не является.
        /// </summary>
        public static readonly IReadOnlyList<string> StockRowSuffixes = new[] { "тз", "в_пути" };

        public const string SoldRowSuffix = "прод";
    }

    /// <summary>Лист «КС»: доля продаж по неделям для каждой подгруппы.</summary>
    public static class Seasonality
    {
        public const string Sector = "сектор";
        public const string Group = "группа";
        public const string Subgroup = "подгруппа";

        /// <summary>Последняя колонка листа - итог, а не неделя.</summary>
        public const string Total = "итого_%";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Sector, new[] { "сектор" }, exactOnly: true),
            new ColumnSpec(Group, new[] { "группа" }, exactOnly: true),
            new ColumnSpec(Subgroup, new[] { "подгруппа" }, exactOnly: true),
        };
    }

    /// <summary>
    /// Лист «Распред». Заголовков в привычном смысле у него нет: разметка задаётся
    /// подписями в колонке слева от блоков и строкой заголовков таблицы РТТ.
    /// </summary>
    public static class Distribution
    {
        /// <summary>Подписи в строке заголовков таблицы РТТ. По ним она и находится.</summary>
        public const string BlockDistribute = "распред";

        public const string BlockHubStock = "запас на хаб";
        public const string BlockStock = "остатки";
        public const string BlockShip = "в загрузку";

        /// <summary>Подписи в колонке слева от блоков.</summary>
        public const string LabelCode = "код";

        public const string LabelUnits = "ед";
        public const string LabelRemainder = "ост";
        public const string LabelHubMinimum = "мин запас на хаб";

        /// <summary>Подпись слева от колонок месяцев.</summary>
        public const string SeasonalityLabel = "сезонность по нк";

        /// <summary>Ширина блока одного АЦР: «распред», «запас на ХАБ», «Остатки», «в загрузку».</summary>
        public const int BlockWidth = 4;

        /// <summary>Ниже этой доли остатка «мин запас на Хаб» обнуляется.</summary>
        public const double MinimumRemainderShare = 0.30;
    }

    /// <summary>
    /// Лист «Загрузочник». Заголовки у него есть, но нужны всего две подписи: по ним
    /// находится строка заголовков, а между ними лежат колонки кодов поставки.
    /// </summary>
    public static class Loader
    {
        public const string LabelCode = "код";

        public const string LabelTotal = "итого";
    }

    /// <summary>Лист инвойса. Нужны только четыре колонки - иначе труднее найти строку заголовков.</summary>
    public static class Invoice
    {
        public const string Barcode = "штрих-код";
        public const string Quantity = "количество";
        public const string Price = "цена";
        public const string Amount = "сумма";

        /// <summary>Подпись в шапке инвойса, под которой стоит номер поставки.</summary>
        public const string SupplyLabel = "инвойс №";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Barcode, new[] { "штрих-код", "штрихкод", "barcode" }, exactOnly: true),
            new ColumnSpec(Quantity, new[] { "количество", "quantity" }, exactOnly: true),
            new ColumnSpec(Price, new[] { "цена", "price" }, exactOnly: true),
            new ColumnSpec(Amount, new[] { "сумма", "amount" }, exactOnly: true),
        };
    }

    /// <summary>Лист «для цен»: подписи в первой колонке, значения во второй и далее.</summary>
    public static class ForPrices
    {
        public const string Supply = "поставка";
        public const string Quantity = "количество";
        public const string Amount = "сумма закупа";
    }

    /// <summary>
    /// Лист «Цены». Перечислены только колонки, которые программа читает или пишет.
    /// Почти все ищутся полным совпадением: на листе рядом стоят «Всп» и «ВспР»,
    /// три колонки «проверка», «Группа» и «Подгруппа», «Острова» и «Период продаж острова».
    /// </summary>
    public static class Prices
    {
        public const string PurchasePrice = "цена_закупочная";
        public const string Supply = "Поставка";
        public const string BarcodeCopy = "Копия ШК";
        public const string Code = "КОД";
        public const string Model = "модель";
        public const string Sector = "Сектор";
        public const string Group = "Группа";
        public const string Subgroup = "Подгруппа";
        public const string Name = "наименование";
        public const string Vsp = "Всп";
        public const string VspR = "ВспР";
        public const string Article = "артикул";
        public const string Color = "цвет";
        public const string Season = "сезон";
        public const string Size = "размер";
        public const string Units = "Ед.";
        public const string MaxPurchase = "Мах_Закуп";
        public const string PurchaseAmount = "Сумма_закупа";
        /// <summary>Ссылается на лист «для цен». Для второй поставки колонка в ссылке другая.</summary>
        public const string LogisticsPercent = "% логистики";

        public const string ActualRate = "Фактический курс";
        public const string CommentKm = "Ком КМ";
        public const string CommentBuyer = "Ком Баера";
        public const string Decision = "Решение";
        public const string PlannedMarkup = "Согласованная наценка по поставке";
        public const string MinMarkup = "Согласованная мин наценка на арт";
        public const string Islands = "Острова";
        public const string StoreSalesPeriod = "Период продаж магазины";
        public const string IslandSalesPeriod = "Период продаж острова";
        public const string WallAnalogue = "Аналог стенки";
        public const string GoogleAnalogue = "Аналог гугл";
        public const string PriceCheck = "Проверка цен";
        public const string Link = "Линк";
        public const string BanCheck = "Проверка запретов";

        /// <summary>
        /// Колонки грейдов. В книге это отдельные колонки «A», «B», «C2», «C1», «D», «E», «F», «G»,
        /// в каждой 1, если этот АЦР положен грейду.
        /// </summary>
        public static readonly IReadOnlyList<string> Grades = new[]
        {
            "A", "B", "C2", "C1", "D", "E", "F", "G",
        };

        /// <summary>Остаток маркетплейса из стенок. Если он равен количеству поставки, это может быть поставка МП.</summary>
        public const string Marketplace = "МП";

        /// <summary>Цена из файла согласования. С ней «Проверка цен» сравнивает розничную цену сети.</summary>
        public const string ApprovedPrice = "Согласованные цены";

        /// <summary>
        /// Колонки, без которых лист «Цены» считается и так: их может не оказаться в заготовке
        /// прошлых поставок. Если колонка есть - программа с ней работает, нет - пропускает.
        /// </summary>
        public static readonly IReadOnlyList<ColumnSpec> OptionalSpecs = new[]
        {
            new ColumnSpec(Marketplace, new[] { "мп" }, exactOnly: true),
            new ColumnSpec(ApprovedPrice, new[] { "согласованные цены", "согласованная цена" }, exactOnly: true),
            new ColumnSpec(IslandGrade120, new[] { "o120", "о120" }, exactOnly: true),
            new ColumnSpec(IslandGrade140, new[] { "o140", "о140" }, exactOnly: true),
        };

        /// <summary>
        /// Грейды островов. Колонки есть не во всех заготовках, поэтому они необязательные,
        /// а буква «О» в заголовке бывает и латинской, и кириллической.
        /// </summary>
        public const string IslandGrade120 = "O120";

        public const string IslandGrade140 = "O140";

        public static readonly IReadOnlyList<string> IslandGrades = new[] { IslandGrade120, IslandGrade140 };

        /// <summary>Все грейды, которые сверяются с «link»: магазинов и островов.</summary>
        public static readonly IReadOnlyList<string> AllGrades = new[]
        {
            "A", "B", "C2", "C1", "D", "E", "F", "G", IslandGrade120, IslandGrade140,
        };

        public static readonly IReadOnlyList<ColumnSpec> Specs = BuildSpecs();

        private static IReadOnlyList<ColumnSpec> BuildSpecs()
        {
            var specs = new List<ColumnSpec>
            {
                new(PurchasePrice, new[] { "цена_закупочная", "цена закупочная" }, exactOnly: true),
                new(Supply, new[] { "поставка" }, exactOnly: true),
                new(BarcodeCopy, new[] { "копия шк" }, exactOnly: true),
                new(Code, new[] { "код" }, exactOnly: true),
                new(Model, new[] { "модель" }, exactOnly: true),
                new(Sector, new[] { "сектор" }, exactOnly: true),
                new(Group, new[] { "группа" }, exactOnly: true),
                new(Subgroup, new[] { "подгруппа" }, exactOnly: true),
                new(Name, new[] { "наименование" }, exactOnly: true),

                // «ВспР» ищется раньше «Всп»: у них общее начало, и колонка,
                // занятая первым совпадением, второму уже не достанется.
                new(VspR, new[] { "вспр", "ацр" }, exactOnly: true),
                new(Vsp, new[] { "всп", "ац" }, exactOnly: true),

                new(Article, new[] { "артикул" }, exactOnly: true),
                new(Color, new[] { "цвет" }, exactOnly: true),
                new(Season, new[] { "сезон" }, exactOnly: true),
                new(Size, new[] { "размер" }, exactOnly: true),
                new(Units, new[] { "ед.", "ед" }, exactOnly: true),
                new(MaxPurchase, new[] { "мах_закуп", "макс_закуп" }, exactOnly: true),
                new(PurchaseAmount, new[] { "сумма_закупа" }, exactOnly: true),
                new(LogisticsPercent, new[] { "% логистики" }, exactOnly: true),
                new(ActualRate, new[] { "фактический курс" }, exactOnly: true),
                new(CommentKm, new[] { "ком км" }, exactOnly: true),
                new(CommentBuyer, new[] { "ком баера" }, exactOnly: true),
                new(Decision, new[] { "решение" }, exactOnly: true),
                new(PlannedMarkup, new[] { "согласованная наценка по поставке" }, exactOnly: true),
                new(MinMarkup, new[] { "согласованная мин наценка на арт" }, exactOnly: true),
                new(Islands, new[] { "острова" }, exactOnly: true),
                new(StoreSalesPeriod, new[] { "период продаж магазины" }, exactOnly: true),
                new(IslandSalesPeriod, new[] { "период продаж острова" }, exactOnly: true),
                new(WallAnalogue, new[] { "аналог стенки" }, exactOnly: true),
                new(GoogleAnalogue, new[] { "аналог гугл" }, exactOnly: true),
                new(PriceCheck, new[] { "проверка цен", "проверка цены" }, exactOnly: true),
                new(Link, new[] { "линк" }, exactOnly: true),
                new(BanCheck, new[] { "проверка запретов" }, exactOnly: true),
            };

            specs.AddRange(Grades.Select(grade =>
                new ColumnSpec(grade, new[] { grade.ToLowerInvariant() }, exactOnly: true)));

            return specs;
        }
    }

    /// <summary>Лист «Сводный прайс»: справочник по штрихкоду.</summary>
    public static class SummaryPrice
    {
        public const string Barcode = "Штрихкод";
        public const string Code = "Код";
        public const string SupplierArticle = "Артикул поставщика";
        public const string Sector = "Сектор";
        public const string Group = "Группа";
        public const string Name = "Наименование";
        public const string Article = "Артикул";
        public const string Color = "Цвет";
        public const string Season = "Сезон";
        public const string Size = "Размер";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Barcode, new[] { "штрихкод", "штрих-код" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),

            // «Артикул поставщика» ищется раньше «Артикула»: иначе более короткое
            // название заняло бы не ту колонку.
            new ColumnSpec(SupplierArticle, new[] { "артикул поставщика", "артикулпоставщика" }),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),

            new ColumnSpec(Sector, new[] { "сектор" }, exactOnly: true),
            new ColumnSpec(Group, new[] { "группа" }, exactOnly: true),
            new ColumnSpec(Name, new[] { "наименование" }, exactOnly: true),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(Season, new[] { "сезон" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),
        };
    }

    /// <summary>Лист «Прайс по подразделениям»: подгруппа по коду.</summary>
    public static class DivisionPrice
    {
        public const string Code = "Код";
        public const string Subgroup = "Подгруппа";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(Subgroup, new[] { "подгруппа" }, exactOnly: true),
        };
    }

    /// <summary>Лист «Регламент наценок».</summary>
    public static class Markup
    {
        public const string Sector = "Сектор";
        public const string SectorPlus = "Сектор+";
        public const string Planned = "плановая наценка";
        public const string Minimum = "Наценка мин/контроль";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Sector, new[] { "сектор" }, exactOnly: true),
            new ColumnSpec(SectorPlus, new[] { "сектор+" }, exactOnly: true),

            // Заголовки длинные («плановая наценка на сезон/контроль по каждой поставке»),
            // поэтому сопоставляются началом строки. «Наценка мин/контроль» отличается
            // от «Наценка мин. для тестовых артикулов» уже четвёртым словом.
            new ColumnSpec(Planned, new[] { "плановая наценка" }),
            new ColumnSpec(Minimum, new[] { "наценка мин/контроль", "наценка мин/контр" }),
        };
    }

    /// <summary>
    /// Лист «Аналоги»: слева АЦ, справа его аналог - «последний АЦ, который приходит в сеть».
    /// Колонки «ВПР», которые вставляют между ними вручную, программе не нужны.
    /// </summary>
    public static class Analogues
    {
        public const string Vsp = "АЦ";
        public const string Analogue = "Аналог";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Vsp, new[] { "ац", "всп" }, exactOnly: true),
            new ColumnSpec(Analogue, new[] { "аналог" }),
        };
    }

    /// <summary>Лист «link»: по строке на каждый магазин для каждого АЦР.</summary>
    public static class Link
    {
        public const string Acr = "Ацр";
        public const string DenyGoodsDivision = "Deny goods подразделение";
        public const string DenyGoodsAll = "Deny goods все подразделения";
        public const string Included = "Вкл";
        /// <summary>Грейд магазина строки.</summary>
        public const string StoreGrade = "Группа_ам";
        public const string DateStart = "Датаначало";
        public const string DateEnd = "Датаокончание";

        /// <summary>
        /// «Концепт» магазина. Необязательная колонка: по ней узнаются острова, у которых
        /// «Группа_ам» пустая, - так в «link» бывает.
        /// </summary>
        public const string Concept = "Концепт";

        /// <summary>Часть «Концепта», по которой строка считается островом.</summary>
        public const string IslandConceptMarker = "остров";

        /// <summary>Группа островов, у которых в «link» нет грейда O120 / O140.</summary>
        public const string IslandsWithoutGrade = "острова без «Группа_ам»";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Acr, new[] { "ацр", "вспр" }, exactOnly: true),
            new ColumnSpec(DenyGoodsDivision, new[] { "deny goods подразделение", "denny goods подразделение" }),
            new ColumnSpec(DenyGoodsAll, new[] { "deny goods все", "denny goods все" }),
            new ColumnSpec(Included, new[] { "вкл" }, exactOnly: true),
            new ColumnSpec(StoreGrade, new[] { "группа_ам", "группа ам" }, exactOnly: true),
            new ColumnSpec(DateStart, new[] { "датаначало", "дата начала", "дата начало" }),
            new ColumnSpec(DateEnd, new[] { "датаокончание", "дата окончания", "датаокончани" }),
        };
    }
}
