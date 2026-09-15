using System.Security.Cryptography;
using WarehousePlanAutomation.Core.Updates;

// Подпись выпусков программы.
//
//   keygen          - создать ключ (один раз); печатает открытый ключ для UpdateSignature.PublicKey
//   pubkey          - напечатать открытый ключ
//   sign <файл.exe> - положить рядом «файл.exe.sig»
//   verify <файл.exe> - проверить подпись встроенным в программу открытым ключом
//
// Закрытый ключ создаётся в хранилище ключей Windows текущего пользователя и не выгружается:
// скопировать его файлом нельзя. Если ключ потерян (переустановка Windows, другой компьютер),
// создаётся новый, открытый ключ в программе меняется, и следующую версию пользователи один раз
// скачивают с GitHub вручную - старая программа новую подпись не примет.

const string KeyName = "WarehousePlanAutomation.ReleaseSigning";

if (args.Length == 0)
{
    return Usage();
}

switch (args[0])
{
    case "keygen":
    {
        if (CngKey.Exists(KeyName))
        {
            Console.Error.WriteLine("Ключ уже есть. Новый создаётся только если старый потерян: удалите его вручную.");
            return 1;
        }

        using var key = CngKey.Create(CngAlgorithm.ECDsaP256, KeyName, new CngKeyCreationParameters
        {
            ExportPolicy = CngExportPolicies.None,
            KeyUsage = CngKeyUsages.Signing,
        });
        using var ecdsa = new ECDsaCng(key);
        Console.WriteLine(Convert.ToBase64String(ecdsa.ExportSubjectPublicKeyInfo()));
        return 0;
    }

    case "pubkey":
    {
        using var ecdsa = OpenKey();
        Console.WriteLine(Convert.ToBase64String(ecdsa.ExportSubjectPublicKeyInfo()));
        return 0;
    }

    case "sign" when args.Length == 2:
    {
        var path = Path.GetFullPath(args[1]);
        var name = Path.GetFileName(path);
        using var ecdsa = OpenKey();
        byte[] digest;
        using (var stream = File.OpenRead(path))
        {
            digest = UpdateSignature.Digest(name, stream);
        }

        var signature = Convert.ToBase64String(ecdsa.SignHash(digest));
        File.WriteAllText(path + UpdateSignature.Extension, signature);

        using (var check = File.OpenRead(path))
        {
            var ownKey = Convert.ToBase64String(ecdsa.ExportSubjectPublicKeyInfo());
            if (!UpdateSignature.IsValid(name, check, signature, ownKey))
            {
                Console.Error.WriteLine("Подпись не проверилась собственным ключом.");
                return 2;
            }
        }

        Console.WriteLine("Подписано: " + path + UpdateSignature.Extension);
        return 0;
    }

    case "verify" when args.Length == 2:
    {
        var path = Path.GetFullPath(args[1]);
        using var stream = File.OpenRead(path);
        var ok = UpdateSignature.IsValid(Path.GetFileName(path), stream, File.ReadAllText(path + UpdateSignature.Extension));
        Console.WriteLine(ok ? "Подпись сходится с ключом программы." : "Подпись НЕ сходится с ключом программы.");
        return ok ? 0 : 3;
    }

    default:
        return Usage();
}

static ECDsaCng OpenKey()
{
    if (!CngKey.Exists(KeyName))
    {
        throw new InvalidOperationException("Ключа подписи на этом компьютере нет: сначала «keygen».");
    }

    return new ECDsaCng(CngKey.Open(KeyName));
}

static int Usage()
{
    Console.Error.WriteLine("ReleaseSigner keygen | pubkey | sign <файл.exe> | verify <файл.exe>");
    return 1;
}
