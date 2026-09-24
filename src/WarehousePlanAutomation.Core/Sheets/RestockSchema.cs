using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Text;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>
/// Город подтоварки: короткое название («Мск»), колонка с его количеством и заголовок
/// этой колонки целиком - его программа называет в сообщениях.
/// </summary>
public sealed record RestockCity(string Name, int Column, string Title = "");

/// <summary>
/// Колонка, по которой считается подтоварка, и города. Городов нет - считаем по одной
/// колонке «в подтоварку»; города есть - по итоговой, а по городам делятся только
/// листы «из‹место›» и загрузочники.
/// </summary>
public sealed record RestockQuantityLayout(int Total, IReadOnlyList<RestockCity> Cities);

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

    /// <summary>
    /// Лист аналитика: секторы, группы и АЦР, которые отдаём без согласования, что бы
    /// ни стояло в «Запрете». Листа нет - исключений нет.
    /// </summary>
    public const string ExceptionsSheet = "Исключения";

    /// <summary>Заказы, которые по «коменту» грузятся отдельно.</summary>
    public const string SeparateSheet = "Отдельно";

    /// <summary>«комент» строки, которую грузят отдельным заказом: «ОЧКИ - отдельный заказ».</summary>
    public const string SeparateMarker = "отдельн";

    public const string NoteOk = "Ок";

    public const string NoteApprove = "Согласовать";

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

        /// <summary>Решение по «Запрету». Программа дописывает колонку, если её нет.</summary>
        public const string Note = "Заметка";

        /// <summary>
        /// Колонка аналитика рядом с «Заметками». Программа её только создаёт и оставляет
        /// пустой: что в ней считать, решает аналитик.
        /// </summary>
        public const string Check = "Проверка";

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

            // «Прогнозный sellout» на решение по строке больше не влияет и нужен только
            // на листе согласования - без него книга всё равно разбирается.
            new ColumnSpec(Sellout, new[] { "прогнозный sellout", "sellout" }, optional: true),
            new ColumnSpec(Note, new[] { "заметка", "заметки" }, exactOnly: true, optional: true),
            new ColumnSpec(Check, new[] { "проверка" }, exactOnly: true, optional: true),
        };

        /// <summary>
        /// Колонки, которые программа дописывает сама, если их нет в книге. Они идут в этом
        /// порядке сразу за колонками аналитика, а дальше - «Место хранения» и остатки мест.
        /// </summary>
        public static readonly IReadOnlyList<CreatedColumn> Created = new[]
        {
            new CreatedColumn(Note, "Заметки", new[] { "заметка", "заметки" }),
            new CreatedColumn(Check, Check, new[] { "проверка" }),
        };

        /// <summary>Итог по всем городам: по нему идут все расчёты подтоварки.</summary>
        public static readonly IReadOnlyList<string> TotalKeys = new[]
        {
            "итого в подтоварку", "итого на мп",
        };

        /// <summary>
        /// Заголовки листа «на загрузку». Колонка количества ищется отдельно от остальных:
        /// городов может быть несколько.
        /// </summary>
        public static HeaderMap ResolveHeaders(SheetGrid grid)
        {
            var headers = HeaderResolver.Resolve(grid, LoadSheet, Specs);
            var quantity = ResolveQuantity(grid, headers.HeaderRow);

            var columns = new Dictionary<string, int>(headers.Columns, StringComparer.Ordinal)
            {
                [Quantity] = quantity.Total,
            };

            return new HeaderMap(headers.HeaderRow, columns);
        }

        /// <summary>Города подтоварки в порядке колонок. Пусто - город один.</summary>
        public static IReadOnlyList<RestockCity> ResolveCities(SheetGrid grid, int headerRow) =>
            ResolveQuantity(grid, headerRow).Cities;

        /// <summary>
        /// Колонка количества и города.
        ///
        /// Один город - одна колонка «в подтоварку»; она же может быть подписана складом
        /// («в подтоварку Мск»). Городов несколько - на каждый своя колонка («в подтоварку Мск»
        /// или «Склад Мск») плюс итоговая «Итого в подтоварку» («Итого на МП»): все расчёты
        /// идут по итогу, а по городам делятся только листы «из‹место›» и загрузочники.
        /// </summary>
        public static RestockQuantityLayout ResolveQuantity(SheetGrid grid, int headerRow)
        {
            var quantityKey = TextUtils.NormalizeKey(Quantity);
            const string warehouseKey = "склад";

            var named = new List<RestockCity>();
            var warehouses = new List<RestockCity>();
            var plain = new List<RestockCity>();
            int? total = null;

            for (var column = grid.FirstColumn; column <= grid.LastColumn; column++)
            {
                var title = TextUtils.Normalize(grid.Text(headerRow, column));
                var key = TextUtils.NormalizeKey(title);
                if (key.Length == 0)
                {
                    continue;
                }

                if (total is null && TotalKeys.Any(name => key.StartsWith(name, StringComparison.Ordinal)))
                {
                    total = column;
                    continue;
                }

                if (key.StartsWith(quantityKey, StringComparison.Ordinal))
                {
                    var city = title[quantityKey.Length..].Trim();
                    (city.Length == 0 ? plain : named).Add(
                        new RestockCity(city.Length == 0 ? title : city, column, title));
                }
                else if (key.StartsWith(warehouseKey + " ", StringComparison.Ordinal))
                {
                    warehouses.Add(new RestockCity(title[(warehouseKey.Length + 1)..].Trim(), column, title));
                }
            }

            // Городские колонки - те, которых больше одной: «в подтоварку Мск» и «… Спб»
            // либо «Склад Мск» и «Склад Спб». Одинокий «Склад …» городом не считается:
            // в книге хватает колонок с похожим началом.
            var cities = named.Count > 1 ? named : warehouses.Count > 1 ? warehouses : new List<RestockCity>();
            var single = plain.Concat(named).OrderBy(city => city.Column).ToList();

            // Итог есть - считаем по нему; города при нём нужны только для деления
            // загрузочников.
            if (total is { } totalColumn)
            {
                return new RestockQuantityLayout(totalColumn, cities);
            }

            if (cities.Count > 1)
            {
                throw new WorkbookValidationException(new[]
                {
                    "на листе «" + LoadSheet + "» несколько колонок с количеством по городам: " +
                    string.Join(", ", cities.Select(city => "«" + city.Title + "»")) +
                    ", а итоговой колонки «Итого в подтоварку» нет. Подтоварка на несколько городов " +
                    "считается по итогу - добавьте его или оставьте один город",
                });
            }

            if (single.Count == 1)
            {
                return new RestockQuantityLayout(single[0].Column, Array.Empty<RestockCity>());
            }

            if (single.Count > 1)
            {
                throw new WorkbookValidationException(new[]
                {
                    "на листе «" + LoadSheet + "» несколько колонок «" + Quantity + "»: " +
                    string.Join(", ", single.Select(city => "«" + city.Title + "»")) +
                    ". Непонятно, по какой считать - оставьте одну или добавьте «Итого в подтоварку»",
                });
            }

            throw new WorkbookValidationException(new[]
            {
                "на листе «" + LoadSheet + "» не найдена колонка «" + Quantity + "»",
            });
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
