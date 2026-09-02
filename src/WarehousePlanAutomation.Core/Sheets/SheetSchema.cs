namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Названия листов и колонок, зафиксированные техническим заданием.
/// Все имена ищутся динамически, буквы колонок и номера строк нигде не «зашиты».
/// </summary>
public static class SheetSchema
{
    public const string PlanSheet = "План";
    public const string OrdersSheet = "Заказы на отгрузку";
    public const string JournalSheet = "Журнал заказов на отгрузку";

    public static class Orders
    {
        public const string Division = "Подразделение";
        public const string Comment = "Комментарий";
        public const string DifferenceUnits = "разница единиц";
        public const string DocumentDate = "Дата документа";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Division, new[] { "подразделение" }),
            new ColumnSpec(Comment, new[] { "комментарий", "комментарии" }),
            new ColumnSpec(DifferenceUnits, new[] { "разница единиц", "разница ед", "разница" }),
            new ColumnSpec(DocumentDate, new[] { "дата документа" }),
        };
    }

    public static class Journal
    {
        public const string Comment = "Комментарий";
        public const string Status = "Статус";
        public const string Percent = "%";

        /// <summary>Номер документа магазина, например «З000-260355». Задаёт единицу подсчёта.</summary>
        public const string DocumentNumber = "Номер";

        /// <summary>Сколько единиц заказано по документу.</summary>
        public const string Quantity = "Кол-во ед";

        /// <summary>Сколько единиц собрано фактически.</summary>
        public const string ActualQuantity = "Факт ед";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Comment, new[] { "комментарий", "комментарии" }),
            new ColumnSpec(Status, new[] { "статус" }),
            new ColumnSpec(Percent, new[] { "%" }, exactOnly: true),
            new ColumnSpec(DocumentNumber, new[] { "номер" }, exactOnly: true),
            new ColumnSpec(Quantity, new[] { "кол-во ед" }, exactOnly: true),
            new ColumnSpec(ActualQuantity, new[] { "факт ед" }, exactOnly: true),
        };
    }

    public static class Plan
    {
        public const string Number = "№";
        public const string Supplies = "Поставки";
        public const string Processing = "Обработка";
        public const string Group = "Группа";
        public const string Comments = "Комментарии";
        public const string DocumentDate = "Дата документа";
        public const string Deadlines = "Сроки выполнения";
        public const string DaysInWork = "Дней в работе";
        public const string LoadNumber = "Номер загрузки";
        public const string Quantity = "Количество единиц";
        public const string Prices = "Цены";
        public const string Priorities = "Приоритеты";
        public const string Status = "статус";
        public const string CompletionPercent = "% выполнения";
        public const string NetworkDate = "Дата в сети";
        public const string ShipmentDate = "Дата отгрузки";
        public const string Decision = "Решение";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Number, new[] { "№", "n" }, exactOnly: true),
            new ColumnSpec(Supplies, new[] { "поставки" }),
            new ColumnSpec(Processing, new[] { "обработка" }),
            new ColumnSpec(Group, new[] { "группа" }),
            new ColumnSpec(Comments, new[] { "комментарии", "комментарий" }),
            new ColumnSpec(DocumentDate, new[] { "дата документа" }),
            new ColumnSpec(Deadlines, new[] { "сроки выполнения" }),
            new ColumnSpec(DaysInWork, new[] { "дней в работе" }),
            new ColumnSpec(LoadNumber, new[] { "номер загрузки" }),
            new ColumnSpec(Quantity, new[] { "количество единиц" }),
            new ColumnSpec(Prices, new[] { "цены" }),
            new ColumnSpec(Priorities, new[] { "приоритеты" }),
            new ColumnSpec(Status, new[] { "статус" }, exactOnly: true),
            new ColumnSpec(CompletionPercent, new[] { "% выполнения" }),
            new ColumnSpec(NetworkDate, new[] { "дата в сети" }),
            new ColumnSpec(ShipmentDate, new[] { "дата отгрузки" }),
            new ColumnSpec(Decision, new[] { "решение" }),
        };

        /// <summary>Названия секций листа «План» (ищутся в колонке «№»).</summary>
        public const string SectionAllGroups = "все группы";

        public const string SectionReturns = "возвраты";
        public const string SectionStorageAcceptance = "приемка на хранилище";
        public const string SectionMarketplaces = "заказы мп";
        public const string SectionWholesale = "заказы опт";
        public const string SectionInternetShop = "заказы им";

        /// <summary>Подстроки блока «заказы МП» (ищутся в колонке «Комментарии»).</summary>
        public const string MarketplaceFromStorage = "с хранилища и хранения";

        public const string MarketplaceFromReturns = "из возвратов";
        public const string MarketplaceFromSupplies = "из поставок";

        /// <summary>Особые строки блока «все группы» (ищутся в колонке «Поставки»).</summary>
        public const string StorageAcceptanceRow = "приемка на хранилище";

        public const string AutoHubRow = "автозаказы для хабов";
    }
}
