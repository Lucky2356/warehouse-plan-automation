namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Книга «Приемка на хранилище». Названия листов и колонок в одном месте: буквы колонок
/// нигде не зашиты, всё ищется по заголовку, а листы - по началу названия.
/// </summary>
public static class ReceivingSchema
{
    public const string StockSheet = "Остатки";
    public const string StorageStockSheet = "Т.Остатки";
    public const string ReservesSheet = "Резервы";
    public const string StorageSheet = "А2, А3";
    public const string MarketplaceSheet = "МП";
    public const string IncomingSheet = "Приходы";
    public const string SummarySheet = "итог";
    public const string CollectedSheet = "Поставки собраны";
    public const string NotCollectedSheet = "Поставки не собраны";
    public const string MarketplaceSuppliesSheet = "Поставки МП";
    public const string AddressesSheet = "адреса";
    public const string ContainersSheet = "тары";
    public const string MarketplaceAddressesSheet = "адреса МП";
    public const string MarketplaceContainersSheet = "тары МП";
    public const string WarehouseSheet = "Склад";
    public const string RemovedFromPlanSheet = "Удалили из плана склада";
    public const string PlanSheet = "План склада";

    /// <summary>В книге лист называется сокращённо: «Непринятый товар в получ постав».</summary>
    public const string NotAcceptedSheet = "Непринятый товар";

    /// <summary>Лист бывает не всегда: «если он есть», говорит инструкция.</summary>
    public const string AgreementSheet = "Согласование количества";

    /// <summary>Лист «Остатки»: выгрузка остатков, из которой собирается «итог».</summary>
    public static class Stock
    {
        public const string Sector = "Сектор";
        public const string Group = "Группа";
        public const string Subgroup = "Подгруппа";
        public const string Name = "Наименование";
        public const string Article = "Артикул";
        public const string Code = "Код";
        public const string Size = "Размер";
        public const string Season = "Сезон";
        public const string Color = "Цвет";
        public const string StorageRemainder = "Остаток хранилище";
        public const string SoldTotal = "Продано итого";

        /// <summary>Колонка есть только в свежей выгрузке: после разбора её уже нет.</summary>
        public const string DenyGoods = "Denny goods";

        public static readonly IReadOnlyList<string> DenyGoodsAliases = new[] { "denny goods", "deny goods" };

        /// <summary>Значение «Denny goods», при котором строка удаляется.</summary>
        public const string DeniedForAll = "все";

        /// <summary>
        /// Одиннадцать колонок, которые остаются на листе: от «Сектор» до «Цвет»,
        /// «Остаток хранилище» и «Продано итого». Всё остальное удаляется.
        /// </summary>
        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Sector, new[] { "сектор" }, exactOnly: true),
            new ColumnSpec(Group, new[] { "группа" }, exactOnly: true),
            new ColumnSpec(Subgroup, new[] { "подгруппа" }, exactOnly: true),
            new ColumnSpec(Name, new[] { "наименование" }, exactOnly: true),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),
            new ColumnSpec(Season, new[] { "сезон" }, exactOnly: true),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(
                StorageRemainder,
                new[] { "остаток хранилище", "остаток на хранилище", "остаток хранилища" },
                exactOnly: true),
            new ColumnSpec(SoldTotal, new[] { "продано итого" }, exactOnly: true),
        };
    }

    /// <summary>Лист «Т.Остатки»: остатки хранилища по коду. «Код» - формула от «Артикула».</summary>
    public static class StorageStock
    {
        public const string Code = "Код";
        public const string Article = "Артикул";
        public const string Remainder = "Остатки";

        /// <summary>Буква в коде, которой помечен брак.</summary>
        public const char DefectMark = 'б';

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Remainder, new[] { "остатки" }, exactOnly: true),
        };
    }

    /// <summary>Лист «Резервы».</summary>
    public static class Reserves
    {
        public const string Date = "Дата";
        public const string Type = "Тип";
        public const string Code = "Код";

        /// <summary>Единственный тип, который должен остаться на листе.</summary>
        public const string ShipmentReserve = "РезервОтгрузка";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Date, new[] { "дата" }, exactOnly: true),
            new ColumnSpec(Type, new[] { "тип" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
        };
    }

    /// <summary>Лист «Склад»: выгрузка всех адресов, из неё собираются «А2, А3» и «МП».</summary>
    public static class Warehouse
    {
        public const string Address = "Адрес";
        public const string Container = "Тара";
        public const string Article = "Артикул";
        public const string Code = "Код";
        public const string Sector = "Сектор";
        public const string Group = "Группа";
        public const string Subgroup = "Подгруппа";
        public const string Quality = "Качество";
        public const string Quantity = "Количество";
        public const string Size = "Размер";
        public const string Season = "Сезон";
        public const string Theme = "Тема";
        public const string Color = "Цвет";
        public const string StorageType = "Тип хранения";

        /// <summary>Тип хранения, который идёт на лист «А2, А3».</summary>
        public const string StorageTypeStorage = "Хранение";

        /// <summary>Тип хранения, который идёт на лист «МП».</summary>
        public const string StorageTypeMarketplace = "Маркетплейс";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Address, new[] { "адрес" }, exactOnly: true),
            new ColumnSpec(Container, new[] { "тара" }, exactOnly: true),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(Sector, new[] { "сектор" }, exactOnly: true),
            new ColumnSpec(Group, new[] { "группа" }, exactOnly: true),
            new ColumnSpec(Subgroup, new[] { "подгруппа" }, exactOnly: true),
            new ColumnSpec(Quality, new[] { "качество" }, exactOnly: true),
            new ColumnSpec(Quantity, new[] { "количество" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),
            new ColumnSpec(Season, new[] { "сезон" }, exactOnly: true),
            new ColumnSpec(Theme, new[] { "тема" }, exactOnly: true),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(StorageType, new[] { "типхранения", "тип хранения" }, exactOnly: true),
        };
    }

    /// <summary>
    /// Листы «А2, А3» и «МП»: адреса одного типа хранения. Колонки до «Цвет» заполняются
    /// со «Склада», дальше идут формулы: «В приемку», «Запас нарастающим», «Учитывать»,
    /// «Номер адреса», «Номер тары».
    /// </summary>
    public static class Storage
    {
        public const string Address = "Адрес";
        public const string Container = "Тара";
        public const string Article = "Артикул";
        public const string Code = "Код";
        public const string Sector = "Сектор";
        public const string Group = "Группа";
        public const string Subgroup = "Подгруппа";
        public const string Name = "наименование";
        public const string Quality = "Качество";
        public const string Quantity = "Количество";
        public const string Size = "Размер";
        public const string Season = "Сезон";
        public const string Theme = "Тема";
        public const string Color = "Цвет";
        public const string ToReceive = "В приемку";
        public const string Counted = "Учитывать";
        public const string AddressNumber = "Номер адреса";
        public const string ContainerNumber = "Номер тары";

        /// <summary>
        /// Что куда переносится со «Склада» - по названиям, а не по порядку: на «Складе»
        /// между «Количеством» и «Размером» стоит «Бренд», и вставка подряд сдвигает
        /// «Цвет» в «В приемку». «наименования» на «Складе» нет - колонка остаётся пустой.
        /// </summary>
        public static readonly IReadOnlyList<(string Target, string? Source)> FromWarehouse =
            new (string Target, string? Source)[]
            {
                (Address, Warehouse.Address),
                (Container, Warehouse.Container),
                (Article, Warehouse.Article),
                (Code, Warehouse.Code),
                (Sector, Warehouse.Sector),
                (Group, Warehouse.Group),
                (Subgroup, Warehouse.Subgroup),
                (Name, null),
                (Quality, Warehouse.Quality),
                (Quantity, Warehouse.Quantity),
                (Size, Warehouse.Size),
                (Season, Warehouse.Season),
                (Theme, Warehouse.Theme),
                (Color, Warehouse.Color),
            };

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Address, new[] { "адрес" }, exactOnly: true),
            new ColumnSpec(Container, new[] { "тара" }, exactOnly: true),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(Sector, new[] { "сектор" }, exactOnly: true),
            new ColumnSpec(Group, new[] { "группа" }, exactOnly: true),
            new ColumnSpec(Subgroup, new[] { "подгруппа" }, exactOnly: true),
            new ColumnSpec(Name, new[] { "наименование" }, exactOnly: true),
            new ColumnSpec(Quality, new[] { "качество" }, exactOnly: true),
            new ColumnSpec(Quantity, new[] { "количество" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),
            new ColumnSpec(Season, new[] { "сезон" }, exactOnly: true),
            new ColumnSpec(Theme, new[] { "тема" }, exactOnly: true),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(ToReceive, new[] { "в приемку" }, exactOnly: true),
            new ColumnSpec(Counted, new[] { "учитывать" }, exactOnly: true),
            new ColumnSpec(AddressNumber, new[] { "номер адреса" }, exactOnly: true),
            new ColumnSpec(ContainerNumber, new[] { "номер тары" }, exactOnly: true),
        };
    }

    /// <summary>Лист «Приходы»: поставки, которые разбираются по цветам.</summary>
    public static class Incoming
    {
        public const string Number = "Номер";
        public const string Status = "Статус";
        public const string Difference = "Разница ед";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Number, new[] { "номер" }, exactOnly: true),
            new ColumnSpec(Status, new[] { "статус" }, exactOnly: true),
            new ColumnSpec(Difference, new[] { "разница ед", "разница ед." }, exactOnly: true),
        };
    }

    /// <summary>
    /// Листы «Поставки собраны», «Поставки не собраны», «Поставки МП». Первые две колонки -
    /// формулы («Код» и «Тара/код поставщика»), дальше - строки «Непринятого товара».
    /// </summary>
    public static class Supplies
    {
        public const string Code = "Код";
        public const string ContainerCode = "Тара/код поставщика";
        public const string SupplyNumber = "Номер поставки";
        public const string Barcode = "Шк";

        /// <summary>Сколько колонок в начале листа занято формулами.</summary>
        public const int FormulaColumns = 2;

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(ContainerCode, new[] { "тара/код поставщика" }, exactOnly: true),
            new ColumnSpec(SupplyNumber, new[] { "номер поставки" }, exactOnly: true),
            new ColumnSpec(Barcode, new[] { "шк", "штрихкод" }, exactOnly: true),
        };
    }

    /// <summary>Лист «Непринятый товар в полученных поставках».</summary>
    public static class NotAccepted
    {
        public const string SupplyNumber = "Номер поставки";
        public const string Code = "Код";

        /// <summary>По этой колонке «итог» считает, сколько товара в поставке по коду.</summary>
        public const string Deviation = "Отклонение";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(SupplyNumber, new[] { "номер поставки" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(Deviation, new[] { "отклонение" }, exactOnly: true),
        };
    }

    /// <summary>Лист «План склада»: номера поставок ищутся в колонке «Поставки».</summary>
    public static class Plan
    {
        public const string Supplies = "Поставки";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Supplies, new[] { "поставки" }, exactOnly: true),
        };
    }

    /// <summary>Лист «Согласование количества».</summary>
    public static class Agreement
    {
        public const string Code = "Код";
        public const string AnswerMarketplace = "Ответ МП";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(AnswerMarketplace, new[] { "ответ мп" }, exactOnly: true),
        };
    }

    /// <summary>Лист «итог».</summary>
    public static class Summary
    {
        public const string Sector = Stock.Sector;
        public const string Group = Stock.Group;
        public const string Subgroup = Stock.Subgroup;
        public const string Name = Stock.Name;
        public const string Article = Stock.Article;
        public const string Code = Stock.Code;
        public const string Color = Stock.Color;
        public const string Size = Stock.Size;
        public const string Season = Stock.Season;
        public const string StorageRemainder = Stock.StorageRemainder;
        public const string SoldTotal = Stock.SoldTotal;

        /// <summary>Над первой из этих колонок стоит номер текущей недели.</summary>
        public const string SeasonSharePlan = "доля сезона план";

        public const string Restock = "Допоставить";
        public const string Quantity = "Количество";
        public const string QuantityMarketplace = "Количество МП";
        public const string Place = "Место хранения";
        public const string ContainerCode = "Тара/код поставщика";
        public const string Barcode = "Штрихкод (для поставок)";
        public const string StorageQuantity = "А2, А3";
        public const string MarketplaceQuantity = "МП";
        public const string Collected = "Поставки собраны";
        public const string NotCollected = "Поставки не собраны";
        public const string MarketplaceSupplies = "Поставки МП";

        /// <summary>Колонки, которые переносятся с листа «Остатки» значениями.</summary>
        public static readonly IReadOnlyList<string> FromStock = new[]
        {
            Sector, Group, Subgroup, Name, Article, Code, Color, Size, Season, StorageRemainder, SoldTotal,
        };

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Sector, new[] { "сектор" }, exactOnly: true),
            new ColumnSpec(Group, new[] { "группа" }, exactOnly: true),
            new ColumnSpec(Subgroup, new[] { "подгруппа" }, exactOnly: true),
            new ColumnSpec(Name, new[] { "наименование" }, exactOnly: true),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),
            new ColumnSpec(Season, new[] { "сезон" }, exactOnly: true),
            new ColumnSpec(
                StorageRemainder,
                new[] { "остаток хранилище", "остаток на хранилище", "остаток хранилища" },
                exactOnly: true),
            new ColumnSpec(SoldTotal, new[] { "продано итого" }, exactOnly: true),
            new ColumnSpec(SeasonSharePlan, new[] { "доля сезона план" }, exactOnly: true),
            new ColumnSpec(Restock, new[] { "допоставить" }, exactOnly: true),
            new ColumnSpec(Quantity, new[] { "количество" }, exactOnly: true),
            new ColumnSpec(QuantityMarketplace, new[] { "количество мп" }, exactOnly: true),
            new ColumnSpec(Place, new[] { "место хранения" }, exactOnly: true),
            new ColumnSpec(ContainerCode, new[] { "тара/код поставщика" }, exactOnly: true),
            new ColumnSpec(Barcode, new[] { "штрихкод (для поставок)", "штрихкод" }, exactOnly: true),
            new ColumnSpec(StorageQuantity, new[] { "а2, а3", "а2,а3", "а2-а3" }, exactOnly: true),
            new ColumnSpec(MarketplaceQuantity, new[] { "мп" }, exactOnly: true),
            new ColumnSpec(Collected, new[] { "поставки собраны" }, exactOnly: true),
            new ColumnSpec(NotCollected, new[] { "поставки не собраны" }, exactOnly: true),
            new ColumnSpec(MarketplaceSupplies, new[] { "поставки мп" }, exactOnly: true),
        };
    }

    /// <summary>Листы «адреса» и «адреса МП».</summary>
    public static class Addresses
    {
        public const string Address = "Адрес";
        public const string Code = "Код";
        public const string AddressNumber = "Номер адреса";
        public const string Joined = "Адреса";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Address, new[] { "адрес" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(AddressNumber, new[] { "номер адреса" }, exactOnly: true),
            new ColumnSpec(Joined, new[] { "адреса" }, exactOnly: true),
        };
    }

    /// <summary>
    /// Листы «тары» и «тары МП». Колонка с перечнем тар в книге подписана «Адреса» -
    /// так она называется в шаблоне.
    /// </summary>
    public static class Containers
    {
        public const string Address = "Адрес";
        public const string Container = "Тара";
        public const string Code = "Код";
        public const string AddressNumber = "Номер адреса";
        public const string ContainerNumber = "Номер тары";
        public const string Joined = "Адреса";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Address, new[] { "адрес" }, exactOnly: true),
            new ColumnSpec(Container, new[] { "тара" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(AddressNumber, new[] { "номер адреса" }, exactOnly: true),
            new ColumnSpec(ContainerNumber, new[] { "номер тары" }, exactOnly: true),
            new ColumnSpec(Joined, new[] { "адреса", "тары" }, exactOnly: true),
        };
    }
}
