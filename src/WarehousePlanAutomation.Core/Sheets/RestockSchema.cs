using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Книга подтоварки маркетплейса. Названия листов и колонок в одном месте:
/// буквы колонок нигде не зашиты, всё ищется по заголовку.
/// </summary>
public static class RestockSchema
{
    public const string LoadSheet = "на загрузку";

    /// <summary>
    /// Лист, который программа создаёт сама: строки, которые нужно согласовать.
    /// Называется так же, как в готовых подтоварках аналитика.
    /// </summary>
    public const string ApprovalSheet = "Согласовать";

    /// <summary>Остатки по адресам и таре: из него собираются листы «МП», «А» и «В».</summary>
    public const string AddressSheet = "по адресам и таре";

    /// <summary>Резервы под заказы на отгрузку.</summary>
    public const string ReservesSheet = "Р";

    /// <summary>Непринятый товар в полученных поставках. Поставки на «С» переезжают на «СЗП».</summary>
    public const string SuppliesSheet = "МПП";

    public const string NotCollectedSheet = "Не собрано";

    /// <summary>Заказы, которые по «коменту» грузятся отдельно.</summary>
    public const string SeparateSheet = "Отдельно";

    /// <summary>«комент» строки, которую грузят отдельным заказом: «ОЧКИ - отдельный заказ».</summary>
    public const string SeparateMarker = "отдельн";

    public const string NoteOk = "Ок";

    public const string NoteApprove = "Согласовать";

    /// <summary>
    /// Секторы, которым согласование не нужно: по решению аналитика мелочь везут
    /// без отдельного разговора. Украшения для волос попали сюда по готовой подтоварке -
    /// там все такие строки помечены «Ок».
    /// </summary>
    public static readonly IReadOnlyList<string> SectorsWithoutApproval = new[]
    {
        "БИЖУТЕРИЯ",
        "МЕЛКИЕ АКСЕССУАРЫ",
        "УКРАШЕНИЯ ДЛЯ ВОЛОС",
    };

    /// <summary>Значения колонки «Запрет». Сравниваются без учёта регистра.</summary>
    public static class Ban
    {
        /// <summary>«Отгрузка в рамках заказа МП, запрет забора из розницы».</summary>
        public const string OrderOnly = "отгрузка в рамках заказа мп";

        /// <summary>«Запрет забора из розницы» - забирать нельзя ничего.</summary>
        public const string NoRetail = "запрет забора из розницы";

        /// <summary>
        /// «доп согл» и «запрет, доп согл»: товар везём только после отдельного разговора.
        /// Слово ищется внутри значения, а не с начала.
        /// </summary>
        public const string ExtraApproval = "доп согл";

        /// <summary>АЦР нет в сезонном файле: ограничений нет.</summary>
        public const string NotFound = "#н/д";
    }

    public static class Load
    {
        public const string ClientCode = "Код клиента";
        public const string Sector = "Сектор";
        public const string Group = "Группа";
        public const string Acr = "АЦР";
        public const string Article = "Артикул";
        public const string Color = "Цвет";
        public const string Size = "Размер";
        public const string Code = "Код";
        public const string Price = "РЦ";
        public const string Quantity = "в подтоварку";
        public const string Comment = "комент";
        public const string Priority = "Приоритет";

        /// <summary>
        /// Колонка запретов. В одних книгах она называется «Запрет», в других -
        /// «Комментарий (на запрет)»; это одна и та же колонка.
        /// </summary>
        public const string Ban = "Запрет";

        /// <summary>«Фактическое кол-во, которое можно собрать на ВБ+Озон».</summary>
        public const string Available = "Фактическое кол-во";

        public const string Sellout = "Прогнозный sellout";
        public const string Note = "Заметка";

        /// <summary>
        /// Колонки, которые переносятся на лист согласования: вся содержательная часть
        /// строки до служебных расчётов.
        /// </summary>
        public static readonly IReadOnlyList<string> ApprovalColumns = new[]
        {
            ClientCode, Sector, Group, Acr, Article, Color, Size, Code, Price,
            Quantity, Comment, Priority, Ban, Available, Sellout, Note,
        };

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(ClientCode, new[] { "код клиента" }, exactOnly: true),
            new ColumnSpec(Sector, new[] { "сектор" }, exactOnly: true),
            new ColumnSpec(Group, new[] { "группа" }, exactOnly: true),
            new ColumnSpec(Acr, new[] { "ацр", "вспр" }, exactOnly: true),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),

            // «Код» ищется после «Код клиента»: у них общее начало, а занятая колонка
            // второму совпадению уже не достанется.
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),

            new ColumnSpec(Price, new[] { "рц" }, exactOnly: true),
            // В книге колонка бывает подписана складом: «в подтоварку Мск».
            new ColumnSpec(Quantity, new[] { "в подтоварку" }, prefixOnly: true),
            // «Запрет» ищется раньше «комента»: у названия «Комментарий (на запрет)»
            // при поиске отбрасывается скобка, и без этого порядка колонку забрал бы
            // «комент» - а «Запрет» не нашёлся бы вовсе.
            new ColumnSpec(
                Ban,
                new[] { "запрет", "комментарий (на запрет)", "коментарий (на запрет)" },
                exactOnly: true),

            new ColumnSpec(Comment, new[] { "комент", "комментарий" }, exactOnly: true),
            new ColumnSpec(Priority, new[] { "приоритет" }, exactOnly: true),
            new ColumnSpec(Available, new[] { "фактическое кол-во", "фактическое количество" }),
            new ColumnSpec(Sellout, new[] { "прогнозный sellout", "sellout" }),
            new ColumnSpec(Note, new[] { "заметка", "заметки" }, exactOnly: true),
        };

        /// <summary>
        /// Заголовки листа «на загрузку». Колонка «в подтоварку» может быть подписана складом -
        /// «в подтоварку Мск». Если таких колонок несколько («в подтоварку» и «в подтоварку Мск»,
        /// «… Мск» и «… Нск»), книга останавливается с объяснением: подтоварка по нескольким
        /// складам пока не разбирается, а молча взять одну из колонок нельзя - количество
        /// другого склада потерялось бы.
        /// </summary>
        public static HeaderMap ResolveHeaders(SheetGrid grid)
        {
            var headers = HeaderResolver.Resolve(grid, LoadSheet, Specs);
            var key = TextUtils.NormalizeKey(Quantity);

            var names = new List<string>();
            for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
            {
                if (TextUtils.NormalizeKey(grid.Text(headers.HeaderRow, column)).StartsWith(key, StringComparison.Ordinal))
                {
                    names.Add(TextUtils.Normalize(grid.Text(headers.HeaderRow, column)));
                }
            }

            if (names.Count > 1)
            {
                throw new WorkbookValidationException(new[]
                {
                    "на листе «" + LoadSheet + "» несколько колонок «" + Quantity + "»: " +
                    string.Join(", ", names.Select(name => "«" + name + "»")) +
                    ". Подтоварка сразу по нескольким складам пока не поддерживается - " +
                    "разберите склады отдельными файлами, оставив в каждом одну такую колонку",
                });
            }

            return headers;
        }

        /// <summary>Колонка, которую заполняет расстановка мест хранения.</summary>
        public const string Place = "Место хранения";

        /// <summary>Цена маркетплейса. Если она есть, в загрузочник идёт она, а не «РЦ».</summary>
        public const string MarketplacePrice = "Цена на МП";

        public const string Reserves = "Р";

        /// <summary>Колонки, которые переносятся на листы «из‹место›», «Не собрано», «Отдельно».</summary>
        public static readonly IReadOnlyList<string> PickColumns = new[]
        {
            ClientCode, Sector, Group, Acr, Article, Color, Size, Code, Price, Quantity, Comment, Priority,
        };
    }

    /// <summary>«по адресам и таре» и собранные из него «МП», «А», «В».</summary>
    public static class Stock
    {
        public const string Acr = "АЦР";
        public const string Article = "Артикул";
        public const string Code = "Код";
        public const string Quantity = "Количество";
        public const string Size = "Размер";
        public const string Color = "Цвет";
        public const string StorageType = "Типхранения";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(Quantity, new[] { "количество" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(StorageType, new[] { "типхранения", "тип хранения" }, exactOnly: true),
        };
    }

    /// <summary>«Р» - резервы.</summary>
    public static class Reserves
    {
        public const string Acr = "АЦР";
        public const string Date = "Дата";
        public const string Division = "Подразделение";
        /// <summary>По ней узнаётся название маркетплейса с незнакомым кодом клиента. Колонка не обязательна.</summary>
        public const string DivisionGroup = "Группа подразделения";
        public const string Article = "Артикул";
        public const string Quantity = "Количество";
        public const string Comment = "Коментарий к заказу";
        public const string Color = "Цвет";
        public const string Size = "Размер";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Date, new[] { "дата" }, exactOnly: true),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Quantity, new[] { "количество" }, exactOnly: true),
            new ColumnSpec(Comment, new[] { "коментарий к заказу", "комментарий к заказу", "комментарий", "коментарий" }),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),
        };
    }

    /// <summary>«МПП» и «СЗП» - непринятый товар в полученных поставках.</summary>
    public static class Supplies
    {
        public const string Acr = "АЦР";
        public const string Number = "Номер поставки";
        public const string Article = "Артикул";
        public const string Code = "Код";
        public const string Deviation = "Отклонение";
        public const string Color = "Цвет";
        public const string Size = "Размер";

        public static readonly IReadOnlyList<ColumnSpec> Specs = new[]
        {
            new ColumnSpec(Number, new[] { "номер поставки", "номер" }, exactOnly: true),
            new ColumnSpec(Article, new[] { "артикул" }, exactOnly: true),
            new ColumnSpec(Code, new[] { "код" }, exactOnly: true),
            new ColumnSpec(Deviation, new[] { "отклонение" }, exactOnly: true),
            new ColumnSpec(Color, new[] { "цвет" }, exactOnly: true),
            new ColumnSpec(Size, new[] { "размер" }, exactOnly: true),
        };
    }
}
