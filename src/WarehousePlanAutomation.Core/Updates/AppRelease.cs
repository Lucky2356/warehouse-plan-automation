using System.Globalization;

namespace WarehousePlanAutomation.Core.Updates;

/// <summary>Выпуск программы на GitHub: версия, описание и готовый .exe.</summary>
/// <param name="SignatureUrl">
/// Файл подписи «‹имя .exe›.sig». null - у выпуска подписи нет, ставить его нельзя.
/// </param>
public sealed record AppRelease(
    Version Version,
    string Tag,
    string Notes,
    DateTimeOffset? Published,
    string AssetName,
    string AssetUrl,
    long AssetSize,
    string? SignatureUrl = null);

/// <summary>Чем закончилась проверка обновлений.</summary>
public sealed record UpdateCheck(AppRelease? Release, string? Problem)
{
    public bool HasUpdate => Release is not null && Problem is null;

    public static UpdateCheck None() => new(null, null);

    public static UpdateCheck Failed(string problem) => new(null, problem);
}

/// <summary>
/// Разбор номера версии. Метка выпуска пишется как «v1.9.0», а версия сборки читается
/// из самой программы, поэтому сравнивать их напрямую как строки нельзя.
/// </summary>
public static class AppVersion
{
    /// <summary>
    /// Версия из метки выпуска. Лишний хвост («v1.9.0-beta.2») отбрасывается: до него
    /// стоит обычная версия, а по хвосту сравнивать нечего.
    /// </summary>
    public static Version? Parse(string? tag)
    {
        if (string.IsNullOrWhiteSpace(tag))
        {
            return null;
        }

        var text = tag.Trim();
        if (text.StartsWith("v", StringComparison.OrdinalIgnoreCase))
        {
            text = text[1..];
        }

        var end = 0;
        var parts = 1;
        while (end < text.Length)
        {
            var symbol = text[end];
            if (char.IsDigit(symbol))
            {
                end++;
                continue;
            }

            if (symbol == '.' && parts < 4 && end + 1 < text.Length && char.IsDigit(text[end + 1]))
            {
                parts++;
                end++;
                continue;
            }

            break;
        }

        return Version.TryParse(Normalize(text[..end]), out var version) ? version : null;
    }

    /// <summary>
    /// Version требует хотя бы двух частей, а метка вида «v2» встречается.
    /// </summary>
    private static string Normalize(string text) =>
        text.Length == 0 ? string.Empty : text.Contains('.') ? text : text + ".0";

    /// <summary>
    /// Сравниваются только первые три числа. Четвёртое - номер сборки, он меняется
    /// от каждой пересборки и обновлением не является.
    /// </summary>
    public static bool IsNewer(Version candidate, Version current) =>
        ThreeParts(candidate) > ThreeParts(current);

    private static Version ThreeParts(Version version) =>
        new(version.Major, version.Minor, Math.Max(version.Build, 0));

    /// <summary>Версия в привычном виде: «1.9.0».</summary>
    public static string Display(Version version) =>
        string.Join(
            ".",
            version.Major.ToString(CultureInfo.InvariantCulture),
            version.Minor.ToString(CultureInfo.InvariantCulture),
            Math.Max(version.Build, 0).ToString(CultureInfo.InvariantCulture));
}
