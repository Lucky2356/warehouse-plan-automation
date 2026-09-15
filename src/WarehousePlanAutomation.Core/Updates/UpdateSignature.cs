using System.Security.Cryptography;
using System.Text;

namespace WarehousePlanAutomation.Core.Updates;

/// <summary>Файл обновления не прошёл проверку подлинности и ставиться не должен.</summary>
public sealed class UpdateSignatureException : Exception
{
    public UpdateSignatureException(string message)
        : base(message)
    {
    }
}

/// <summary>
/// Подпись выпуска. Рядом с «WarehousePlanAutomation-1.20.0-win-x64.exe» в выпуске лежит
/// «WarehousePlanAutomation-1.20.0-win-x64.exe.sig» - подпись ECDSA P-256.
///
/// Закрытый ключ хранится только на компьютере разработчика, в хранилище ключей Windows,
/// и в GitHub не попадает. Поэтому доступа к репозиторию или аккаунту GitHub мало, чтобы
/// подсунуть программе свой файл: без ключа подпись к нему не сделать, а файл без верной
/// подписи программа не ставит.
///
/// Подписывается не только содержимое, но и имя файла - в нём версия. Так старый настоящий
/// файл с его настоящей подписью нельзя выдать за новую версию.
/// </summary>
public static class UpdateSignature
{
    /// <summary>Расширение файла подписи рядом с .exe выпуска.</summary>
    public const string Extension = ".sig";

    /// <summary>
    /// Открытый ключ выпусков (SubjectPublicKeyInfo, base64). Его публиковать можно:
    /// им подпись только проверяется.
    /// </summary>
    public const string PublicKey =
        "MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAESYIDc6OWunFzF1V76SaGuvVJI1+BSBinllwEAadk69WVomkMtTFjgm0sgraCfV7RhlGwZdLD8MhLZXHqEUvyKA==";

    /// <summary>Что именно подписывается: имя файла и SHA-256 его содержимого.</summary>
    public static byte[] Digest(string assetName, Stream content)
    {
        var fileHash = Convert.ToHexString(SHA256.HashData(content));
        var message = "WarehousePlanAutomation|" + assetName + "|" + fileHash;
        return SHA256.HashData(Encoding.UTF8.GetBytes(message));
    }

    /// <summary>
    /// Сходится ли подпись. Текст подписи - base64, лишние пробелы и переводы строк
    /// по краям не мешают. Любая ошибка разбора - это «не сходится», а не исключение.
    /// </summary>
    public static bool IsValid(string assetName, Stream content, string? signatureText, string publicKey = PublicKey)
    {
        if (string.IsNullOrWhiteSpace(signatureText))
        {
            return false;
        }

        byte[] signature;
        try
        {
            signature = Convert.FromBase64String(signatureText.Trim());
        }
        catch (FormatException)
        {
            return false;
        }

        try
        {
            using var key = ECDsa.Create();
            key.ImportSubjectPublicKeyInfo(Convert.FromBase64String(publicKey), out _);
            return key.VerifyHash(Digest(assetName, content), signature);
        }
        catch (CryptographicException)
        {
            return false;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    /// <summary>Проверяет файл на диске; при несовпадении - исключение с понятным текстом.</summary>
    public static void EnsureValid(string path, string assetName, string? signatureText)
    {
        using var stream = File.OpenRead(path);
        if (!IsValid(assetName, stream, signatureText))
        {
            throw new UpdateSignatureException(
                "Подпись файла обновления не сходится: файл изменён или выложен не разработчиком.");
        }
    }
}
