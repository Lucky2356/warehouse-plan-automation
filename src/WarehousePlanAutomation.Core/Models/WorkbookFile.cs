namespace WarehousePlanAutomation.Core.Models;

/// <summary>
/// Форматы книг Excel, которые программа умеет обрабатывать.
/// Один список на всех: его проверяет и обработчик, и перетаскивание файла в окно.
/// </summary>
public static class WorkbookFile
{
    public static readonly IReadOnlyList<string> SupportedExtensions = new[]
    {
        ".xlsx",
        ".xlsm",
        ".xlsb",
        ".xls",
    };

    /// <summary>Фильтр для диалога выбора файла.</summary>
    public static string DialogFilter =>
        "Книги Excel (" + string.Join(";", SupportedExtensions.Select(e => "*" + e)) + ")|" +
        string.Join(";", SupportedExtensions.Select(e => "*" + e)) +
        "|Все файлы (*.*)|*.*";

    public static bool IsSupported(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        var extension = Path.GetExtension(path);
        return SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
    }
}
