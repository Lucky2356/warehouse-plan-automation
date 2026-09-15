using System.IO.Compression;
using System.Xml.Linq;

namespace WarehousePlanAutomation.Core.Sheets;

/// <summary>Что удалось узнать о книге, не открывая Excel.</summary>
/// <param name="Sheets">Названия листов в том порядке, в каком они в книге.</param>
/// <param name="ExternalFiles">Файлы, на которые в книге есть внешние ссылки.</param>
public sealed record WorkbookProbe(
    IReadOnlyList<string> Sheets,
    IReadOnlyList<string> ExternalFiles);

/// <summary>
/// Быстрый осмотр книги без Excel.
///
/// Книга .xlsx - это zip, и список листов лежит в одном маленьком файле внутри.
/// Прочитать его - миллисекунды даже для книги на 76 МБ, тогда как открыть такую книгу
/// в Excel это десятки секунд. Поэтому проверить, всё ли на месте, можно сразу после
/// выбора файла, а не после долгой обработки.
///
/// Старый двоичный формат (.xls) не zip - для него осмотр не делается вовсе.
/// </summary>
public static class WorkbookProbeReader
{
    private static readonly XNamespace Spreadsheet =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly XNamespace Relationships =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>Возвращает null, если книгу так прочитать нельзя.</summary>
    public static WorkbookProbe? Read(string path)
    {
        try
        {
            using var archive = ZipFile.OpenRead(path);

            var sheets = ReadSheetNames(archive);
            if (sheets is null)
            {
                return null;
            }

            return new WorkbookProbe(sheets, ReadExternalFiles(archive));
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException
                                       or UnauthorizedAccessException or NotSupportedException)
        {
            // Не zip, файл занят или к нему нет доступа - осмотр просто не делается.
            return null;
        }
    }

    private static IReadOnlyList<string>? ReadSheetNames(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/workbook.xml");
        if (entry is null)
        {
            return null;
        }

        using var stream = entry.Open();
        var document = XDocument.Load(stream);

        return document.Root?
            .Elements(Spreadsheet + "sheets")
            .Elements(Spreadsheet + "sheet")
            .Select(sheet => (string?)sheet.Attribute("name") ?? string.Empty)
            .Where(name => name.Length > 0)
            .ToList();
    }

    /// <summary>
    /// Внешние ссылки книги. Адреса лежат в связях: Target с TargetMode="External".
    /// </summary>
    private static IReadOnlyList<string> ReadExternalFiles(ZipArchive archive)
    {
        var files = new List<string>();

        foreach (var entry in archive.Entries)
        {
            if (!entry.FullName.StartsWith("xl/externalLinks/_rels/", StringComparison.OrdinalIgnoreCase) ||
                !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            using var stream = entry.Open();
            var document = XDocument.Load(stream);

            var targets = document.Root?
                .Elements(Relationships + "Relationship")
                .Where(link => (string?)link.Attribute("TargetMode") == "External")
                .Select(link => (string?)link.Attribute("Target") ?? string.Empty)
                ?? Enumerable.Empty<string>();

            foreach (var target in targets)
            {
                var file = ToLocalPath(target);
                if (file.Length > 0 && !files.Contains(file, StringComparer.OrdinalIgnoreCase))
                {
                    files.Add(file);
                }
            }
        }

        return files;
    }

    /// <summary>
    /// Адрес связи записан как URI («file:///\\FileServer\...»), иногда с процентными
    /// заменами. Нужен обычный путь - его и проверяют на доступность.
    /// </summary>
    private static string ToLocalPath(string target)
    {
        if (target.Length == 0)
        {
            return string.Empty;
        }

        var text = target;
        if (text.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
        {
            text = text["file:///".Length..];
        }
        else if (text.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
        {
            text = text["file://".Length..];
        }

        try
        {
            text = Uri.UnescapeDataString(text);
        }
        catch (UriFormatException)
        {
            // Оставляем как есть: проверка доступности всё равно скажет своё.
        }

        // Excel пишет имя книги в квадратных скобках: «\\сервер\папка\[книга.xlsx]».
        return text.Replace("[", string.Empty, StringComparison.Ordinal)
                   .Replace("]", string.Empty, StringComparison.Ordinal)
                   .Trim();
    }
}
