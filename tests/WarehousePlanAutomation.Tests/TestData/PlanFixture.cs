using WarehousePlanAutomation.Core.Models;
using WarehousePlanAutomation.Core.Sheets;

namespace WarehousePlanAutomation.Tests.TestData;

/// <summary>
/// Синтетический лист «План», повторяющий структуру реального файла:
/// заголовки в первой строке, блоки «все группы», «возвраты», «приемка на хранилище»,
/// «заказы МП» с тремя подстроками, «заказы Опт» и «заказы ИМ».
/// </summary>
public static class PlanFixture
{
    public const long ExistingUrgentLoadNumber = 55575395;
    public const long ExistingSetLoadNumber = 55591671;
    public const long ExistingMonoLoadNumber = 55591776;
    public const long ExistingReturnLoadNumber = 55553189;
    public const long ExistingStorageLoadNumber = 55377185;
    public const string PlaceholderSupplies = "1079-051 Шапки_отгрузка по готовности";

    public static SheetGrid BuildGrid()
    {
        var rows = new List<object?[]>
        {
            Header(),
            Section("все группы"),
            DataRow(1, "Приемка на хранилище от 24.08", quantity: 4334),
            DataRow(2, "Автозаказы для ХАБов", quantity: 42000),
            DataRow(3, "Срочная подтоварка 28.08_Хранение, хранилище",
                loadNumber: ExistingUrgentLoadNumber, quantity: 19021, status: "в сборке.", percent: 83,
                networkDate: 46273),
            DataRow(4, "1100-026, 028 Обувь СЕТ1_в рознице с 1.10",
                loadNumber: ExistingSetLoadNumber, quantity: 576, percent: 0, networkDate: 46296),
            DataRow(5, "1100-026, 028 Обувь МОНО_в рознице с 1.10",
                loadNumber: ExistingMonoLoadNumber, quantity: 324, percent: 0, networkDate: 46296),
            DataRow(6, PlaceholderSupplies, quantity: 2000, percent: 0, networkDate: 46276,
                processing: "перекр", comments: "Заказы будут загружены 27.08"),
            Section("возвраты"),
            DataRow(1, "Зимняя обувь_ликвиды_с хранилища, возвратов",
                loadNumber: ExistingReturnLoadNumber, quantity: 3610, percent: 0, networkDate: 46296),
            Section("приемка на хранилище"),
            DataRow(1, "Сезонный товар FW26-27 из возвратов, времянки_приемка на хранилище_2 приоритет",
                loadNumber: ExistingStorageLoadNumber, quantity: 5914, status: "в сборке.", percent: 0,
                networkDate: 46268),
            Section("заказы МП"),
            MarketplaceRow("с хранилища и хранения", 2339),
            MarketplaceRow("из возвратов", 5281),
            MarketplaceRow("из поставок", 13191),
            Section("заказы Опт", 4590),
            Section("заказы ИМ", 125),
            new object?[17],
        };

        var values = ToArray(rows);
        var formulas = BuildFormulas(rows.Count, 17);
        return new SheetGrid(1, 1, values, formulas);
    }

    public static PlanLayout BuildLayout() => PlanSheetReader.Read(BuildGrid());

    /// <summary>
    /// Тот же лист, но с пустой строкой внутри блока «все группы»:
    /// строки блока перестают идти подряд.
    /// </summary>
    public static PlanLayout BuildLayoutWithGapInAllGroups()
    {
        var rows = new List<object?[]>
        {
            Header(),
            Section("все группы"),
            DataRow(1, "Приемка на хранилище от 24.08", quantity: 4334),
            DataRow(2, "Автозаказы для ХАБов", quantity: 42000),
            new object?[17],
            DataRow(3, "1100-026, 028 Обувь МОНО_в рознице с 1.10",
                loadNumber: ExistingMonoLoadNumber, quantity: 324, percent: 0, networkDate: 46296),
            DataRow(4, "1100-026, 028 Обувь СЕТ1_в рознице с 1.10",
                loadNumber: ExistingSetLoadNumber, quantity: 576, percent: 0, networkDate: 46296),
            Section("возвраты"),
            DataRow(1, "Зимняя обувь_ликвиды_с хранилища, возвратов",
                loadNumber: ExistingReturnLoadNumber, quantity: 3610, percent: 0, networkDate: 46296),
            Section("приемка на хранилище"),
            DataRow(1, "Сезонный товар FW26-27_приемка на хранилище_2 приоритет",
                loadNumber: ExistingStorageLoadNumber, quantity: 5914, percent: 0, networkDate: 46268),
            Section("заказы МП"),
            MarketplaceRow("с хранилища и хранения", 2339),
            MarketplaceRow("из возвратов", 5281),
            MarketplaceRow("из поставок", 13191),
            Section("заказы Опт", 4590),
            Section("заказы ИМ", 125),
        };

        return PlanSheetReader.Read(new SheetGrid(1, 1, ToArray(rows)));
    }

    private static object?[] Header() => new object?[]
    {
        "№", "Поставки", "Обработка", "Группа", "Комментарий", "Дата документа", "Сроки выполнения",
        "Дней в работе", "Номер загрузки", "Количество единиц", "Цены", "Приоритеты", "статус",
        "% выполнения", "Дата в сети (без целевой даты, отгрузка и сборка 15 дн)",
        "Дата отгрузки (10 дней на доставку)", "Решение",
    };

    private static object?[] Section(string name, double? quantity = null)
    {
        var row = new object?[17];
        row[0] = name;
        row[9] = quantity;
        return row;
    }

    private static object?[] MarketplaceRow(string comment, double quantity)
    {
        var row = new object?[17];
        row[4] = comment;
        row[9] = quantity;
        return row;
    }

    private static object?[] DataRow(
        int number,
        string supplies,
        long? loadNumber = null,
        double? quantity = null,
        string? status = null,
        double? percent = null,
        double? networkDate = null,
        string? processing = null,
        string? comments = null)
    {
        var row = new object?[17];
        row[0] = (double)number;
        row[1] = supplies;
        row[2] = processing;
        row[4] = comments ?? "Заказы загружены.";
        row[5] = 46262d;
        row[8] = loadNumber.HasValue ? (double)loadNumber.Value : null;
        row[9] = quantity;
        row[12] = status;
        row[13] = percent;
        row[14] = networkDate;
        return row;
    }

    private static object?[,] ToArray(IReadOnlyList<object?[]> rows)
    {
        var values = new object?[rows.Count, 17];
        for (var r = 0; r < rows.Count; r++)
        {
            for (var c = 0; c < rows[r].Length && c < 17; c++)
            {
                values[r, c] = rows[r][c];
            }
        }

        return values;
    }

    /// <summary>Итоговые формулы блоков и формула «Дата в сети» в строках данных.</summary>
    private static string?[,] BuildFormulas(int rowCount, int columnCount)
    {
        var formulas = new string?[rowCount, columnCount];

        // Строки Excel: 1 заголовок; 2 «все группы» (данные 3..8); 9 «возвраты» (данные 10);
        // 11 «приемка на хранилище» (данные 12); 13 «заказы МП» (подстроки 14..16);
        // 17 «заказы Опт»; 18 «заказы ИМ».
        // Итог блока «все группы» начинается со строки 5: строки 3 и 4 в сумму не входят.
        formulas[1, 9] = "=SUM(J5:J8)";
        formulas[8, 9] = "=SUM(J10:J10)";
        formulas[10, 9] = "=SUM(J12:J12)";
        formulas[12, 9] = "=SUM(J14:J16)";

        foreach (var dataRowIndex in new[] { 4, 5, 6, 7, 9, 11 })
        {
            var excelRow = dataRowIndex + 1;
            formulas[dataRowIndex, 14] = "=F" + excelRow + "+$O$2";
            formulas[dataRowIndex, 15] = "=O" + excelRow + "-$P$2";
        }

        return formulas;
    }
}
