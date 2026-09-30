using System.Security.Cryptography;

namespace WarehousePlanAutomation.Core.Updates;

/// <summary>Файл обновления не прошёл проверку и ставиться не должен.</summary>
public sealed class UpdateVerificationException : Exception
{
    public UpdateVerificationException(string message)
        : base(message)
    {
    }
}

/// <summary>
/// Проверка скачанного обновления по контрольной сумме SHA-256, которую GitHub сам считает
/// для каждого файла выпуска и отдаёт в поле «digest». Ключей и отдельной подписи не нужно:
/// всё берётся из GitHub. Так битая, недокачанная или подменённая по дороге загрузка
/// не будет установлена.
/// </summary>
public static class UpdateChecksum
{
    private const string Prefix = "sha256:";

    /// <summary>
    /// Сумма из поля «digest» («sha256:‹64 шестнадцатеричных знака›») строчными буквами.
    /// Другой алгоритм или мусор - null: проверить файл не по чему.
    /// </summary>
    public static string? ParseDigest(string? digest)
    {
        if (string.IsNullOrWhiteSpace(digest))
        {
            return null;
        }

        var text = digest.Trim();
        if (!text.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var hex = text[Prefix.Length..];
        return hex.Length == 64 && hex.All(Uri.IsHexDigit) ? hex.ToLowerInvariant() : null;
    }

    public static bool IsValid(Stream content, string? expectedSha256)
    {
        if (string.IsNullOrEmpty(expectedSha256))
        {
            return false;
        }

        var actual = Convert.ToHexString(SHA256.HashData(content));
        return string.Equals(actual, expectedSha256, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Проверяет файл на диске; при несовпадении - исключение с понятным текстом.</summary>
    public static void EnsureValid(string path, string? expectedSha256)
    {
        using var stream = File.OpenRead(path);
        if (!IsValid(stream, expectedSha256))
        {
            throw new UpdateVerificationException(
                "Контрольная сумма файла обновления не совпадает с GitHub: файл повреждён при загрузке.");
        }
    }
}
