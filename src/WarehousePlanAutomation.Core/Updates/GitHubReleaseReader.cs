using System.Globalization;
using System.Text.Json;

namespace WarehousePlanAutomation.Core.Updates;

/// <summary>
/// Разбор ответа GitHub о последнем выпуске.
///
/// Вынесен отдельно от загрузки нарочно: разбор - это то, что может однажды поехать
/// (у выпуска не окажется .exe, метка будет написана иначе), и проверять это нужно
/// тестами, а не запросом в сеть.
/// </summary>
public static class GitHubReleaseReader
{
    /// <summary>Файл выпуска: единственный .exe, готовый к запуску.</summary>
    private const string AssetExtension = ".exe";

    public static AppRelease? Read(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return null;
        }

        JsonElement root;
        try
        {
            using var document = JsonDocument.Parse(json);
            root = document.RootElement.Clone();
        }
        catch (JsonException)
        {
            return null;
        }

        if (root.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        // Черновики и предварительные выпуски пользователю не показываются:
        // их выкладывают, чтобы посмотреть самому, а не чтобы их поставили.
        if (Flag(root, "draft") || Flag(root, "prerelease"))
        {
            return null;
        }

        var tag = Text(root, "tag_name");
        var version = AppVersion.Parse(tag);
        if (version is null)
        {
            return null;
        }

        var asset = FindAsset(root);
        if (asset is null)
        {
            return null;
        }

        return new AppRelease(
            version,
            tag,
            Text(root, "body").Replace("\r\n", "\n").Trim(),
            Time(root, "published_at"),
            asset.Value.Name,
            asset.Value.Url,
            asset.Value.Size,
            FindSignature(root, asset.Value.Name));
    }

    /// <summary>Подпись .exe - файл выпуска с тем же именем и «.sig» на конце.</summary>
    private static string? FindSignature(JsonElement root, string assetName)
    {
        var wanted = assetName + UpdateSignature.Extension;
        foreach (var asset in root.GetProperty("assets").EnumerateArray())
        {
            if (string.Equals(Text(asset, "name"), wanted, StringComparison.OrdinalIgnoreCase))
            {
                var url = Text(asset, "browser_download_url");
                return url.Length > 0 ? url : null;
            }
        }

        return null;
    }

    /// <summary>
    /// Из файлов выпуска берётся .exe. Их и должно быть ровно столько: рабочий процесс
    /// выкладывает единственный самодостаточный файл.
    /// </summary>
    private static (string Name, string Url, long Size)? FindAsset(JsonElement root)
    {
        if (!root.TryGetProperty("assets", out var assets) || assets.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        foreach (var asset in assets.EnumerateArray())
        {
            var name = Text(asset, "name");
            if (!name.EndsWith(AssetExtension, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var url = Text(asset, "browser_download_url");
            if (url.Length == 0)
            {
                continue;
            }

            var size = asset.TryGetProperty("size", out var value) && value.TryGetInt64(out var bytes)
                ? bytes
                : 0L;

            return (name, url, size);
        }

        return null;
    }

    private static string Text(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    private static bool Flag(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.True;

    private static DateTimeOffset? Time(JsonElement element, string name) =>
        DateTimeOffset.TryParse(
            Text(element, name),
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
            out var moment)
            ? moment
            : null;
}
