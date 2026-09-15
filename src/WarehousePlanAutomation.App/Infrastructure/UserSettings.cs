using System.IO;
using System.Text.Json;
using WarehousePlanAutomation.Core.Logging;

namespace WarehousePlanAutomation.App.Infrastructure;

/// <summary>
/// Мелкие настройки, которые программа помнит между запусками: пока это только папка,
/// из которой в прошлый раз брали книгу, - у каждой задачи своя.
///
/// Лежит рядом с журналом работы: программа переносится одним файлом, и место
/// для своих настроек у неё то же самое. Не прочиталось или не записалось - не беда:
/// это удобство, а не данные.
/// </summary>
public sealed class UserSettings
{
    private readonly IAppLogger _logger;
    private readonly string _path;
    private readonly Dictionary<string, string> _values;

    public UserSettings(IAppLogger logger)
        : this(logger, DefaultPath())
    {
    }

    public UserSettings(IAppLogger logger, string path)
    {
        _logger = logger;
        _path = path;
        _values = Read();
    }

    public string? Get(string key) => _values.TryGetValue(key, out var value) ? value : null;

    public void Set(string key, string value)
    {
        if (_values.TryGetValue(key, out var current) && current == value)
        {
            return;
        }

        _values[key] = value;
        Write();
    }

    private Dictionary<string, string> Read()
    {
        try
        {
            if (!File.Exists(_path))
            {
                return new Dictionary<string, string>(StringComparer.Ordinal);
            }

            var json = File.ReadAllText(_path);
            return JsonSerializer.Deserialize<Dictionary<string, string>>(json)
                   ?? new Dictionary<string, string>(StringComparer.Ordinal);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            _logger.Warning("Не удалось прочитать настройки, взяты пустые.", ex);
            return new Dictionary<string, string>(StringComparer.Ordinal);
        }
    }

    private void Write()
    {
        try
        {
            var directory = Path.GetDirectoryName(_path);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(_path, JsonSerializer.Serialize(_values));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _logger.Warning("Не удалось сохранить настройки.", ex);
        }
    }

    private static string DefaultPath() =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WarehousePlanAutomation",
            "settings.json");
}
