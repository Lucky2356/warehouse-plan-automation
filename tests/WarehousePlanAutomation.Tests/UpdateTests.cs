using WarehousePlanAutomation.Core.Updates;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class AppVersionTests
{
    [Theory]
    [InlineData("v1.9.0", "1.9.0")]
    [InlineData("1.9.0", "1.9.0")]
    [InlineData("V2.0.1", "2.0.1")]
    [InlineData("v1.9.0-beta.2", "1.9.0")]
    [InlineData("v2", "2.0")]
    public void МеткаВыпускаЧитаетсяКакВерсия(string tag, string expected)
    {
        Assert.Equal(Version.Parse(expected), AppVersion.Parse(tag));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("release")]
    [InlineData(null)]
    public void НеВерсия_НеРазбирается(string? tag)
    {
        Assert.Null(AppVersion.Parse(tag));
    }

    [Theory]
    [InlineData("1.9.0", "1.8.0", true)]
    [InlineData("1.8.1", "1.8.0", true)]
    [InlineData("2.0.0", "1.9.9", true)]
    [InlineData("1.8.0", "1.8.0", false)]
    [InlineData("1.7.0", "1.8.0", false)]
    public void НовееТолькоПоПервымТрёмЧислам(string candidate, string current, bool expected)
    {
        Assert.Equal(
            expected,
            AppVersion.IsNewer(Version.Parse(candidate), Version.Parse(current)));
    }

    [Fact]
    public void НомерСборкиОбновлениемНеСчитается()
    {
        // Сборка меняется от каждой пересборки, поэтому 1.8.0.42 - это всё та же 1.8.0.
        Assert.False(AppVersion.IsNewer(new Version(1, 8, 0, 42), new Version(1, 8, 0, 0)));
    }

    [Fact]
    public void ВерсияПоказываетсяТремяЧислами()
    {
        Assert.Equal("1.8.0", AppVersion.Display(new Version(1, 8, 0, 42)));
    }
}

public class GitHubReleaseReaderTests
{
    private static string Json(
        string tag = "v1.9.0",
        string assets = """
            [{"name":"WarehousePlanAutomation-1.9.0-win-x64.exe",
              "browser_download_url":"https://example.invalid/app.exe","size":71718761}]
            """,
        bool draft = false,
        bool prerelease = false) =>
        $$"""
        {
          "tag_name": "{{tag}}",
          "name": "{{tag}}",
          "body": "Третий этап распреда.\r\nЛист «Загрузочник».",
          "draft": {{(draft ? "true" : "false")}},
          "prerelease": {{(prerelease ? "true" : "false")}},
          "published_at": "2026-09-09T10:15:00Z",
          "assets": {{assets}}
        }
        """;

    [Fact]
    public void ЧитаетВерсиюОписаниеИФайл()
    {
        var release = GitHubReleaseReader.Read(Json());

        Assert.NotNull(release);
        Assert.Equal(new Version(1, 9, 0), release!.Version);
        Assert.Equal("v1.9.0", release.Tag);
        Assert.Equal("WarehousePlanAutomation-1.9.0-win-x64.exe", release.AssetName);
        Assert.Equal("https://example.invalid/app.exe", release.AssetUrl);
        Assert.Equal(71718761, release.AssetSize);
        Assert.Equal(new DateTimeOffset(2026, 9, 9, 10, 15, 0, TimeSpan.Zero), release.Published);
    }

    [Fact]
    public void ПереводыСтрокПриводятсяКОдномуВиду()
    {
        // «\r\n» из ответа GitHub в WPF рисуется лишним пустым знаком.
        var release = GitHubReleaseReader.Read(Json());

        Assert.Equal("Третий этап распреда.\nЛист «Загрузочник».", release!.Notes);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void ЧерновикиИПредварительныеВыпускиНеПредлагаются(bool draft, bool prerelease)
    {
        Assert.Null(GitHubReleaseReader.Read(Json(draft: draft, prerelease: prerelease)));
    }

    [Fact]
    public void ВыпускБезExe_НеПредлагается()
    {
        // Ставить нечего: рабочий процесс выкладывает единственный .exe, и без него
        // предлагать обновление было бы обманом.
        var json = Json(assets: """[{"name":"исходники.zip","browser_download_url":"https://e.invalid/z","size":10}]""");

        Assert.Null(GitHubReleaseReader.Read(json));
    }

    [Fact]
    public void ВыпускБезФайловВообще_НеПредлагается()
    {
        Assert.Null(GitHubReleaseReader.Read(Json(assets: "[]")));
    }

    [Fact]
    public void МеткаБезВерсии_НеПредлагается()
    {
        Assert.Null(GitHubReleaseReader.Read(Json(tag: "latest")));
    }

    [Theory]
    [InlineData("")]
    [InlineData("не json")]
    [InlineData("[1,2,3]")]
    [InlineData(null)]
    public void МусорВОтвете_НеЛомаетПроверку(string? json)
    {
        Assert.Null(GitHubReleaseReader.Read(json));
    }

    [Fact]
    public void ПодписьExe_НаходитсяПоИмени()
    {
        var json = Json(assets: """
            [{"name":"WarehousePlanAutomation-1.9.0-win-x64.exe.sig","browser_download_url":"https://example.invalid/app.exe.sig","size":88},
             {"name":"WarehousePlanAutomation-1.9.0-win-x64.exe","browser_download_url":"https://example.invalid/app.exe","size":71718761}]
            """);

        var release = GitHubReleaseReader.Read(json);

        Assert.Equal("https://example.invalid/app.exe", release!.AssetUrl);
        Assert.Equal("https://example.invalid/app.exe.sig", release.SignatureUrl);
    }

    [Fact]
    public void БезФайлаПодписи_ПодписьПустая()
    {
        Assert.Null(GitHubReleaseReader.Read(Json())!.SignatureUrl);
    }
}

public class UpdateSignatureTests
{
    private const string Name = "WarehousePlanAutomation-1.20.0-win-x64.exe";

    private static readonly byte[] Content = System.Text.Encoding.UTF8.GetBytes("MZ... исполняемый файл");

    private static (string PublicKey, string Signature) Sign(string name, byte[] content)
    {
        using var key = System.Security.Cryptography.ECDsa.Create(System.Security.Cryptography.ECCurve.NamedCurves.nistP256);
        var digest = UpdateSignature.Digest(name, new MemoryStream(content));
        return (Convert.ToBase64String(key.ExportSubjectPublicKeyInfo()), Convert.ToBase64String(key.SignHash(digest)));
    }

    [Fact]
    public void ВернаяПодпись_Сходится()
    {
        var (publicKey, signature) = Sign(Name, Content);

        Assert.True(UpdateSignature.IsValid(Name, new MemoryStream(Content), " " + signature + "\n", publicKey));
    }

    [Fact]
    public void ИзменённыйФайл_НеСходится()
    {
        var (publicKey, signature) = Sign(Name, Content);
        var changed = Content.ToArray();
        changed[0] ^= 1;

        Assert.False(UpdateSignature.IsValid(Name, new MemoryStream(changed), signature, publicKey));
    }

    [Fact]
    public void СтарыйФайлПодДругимИменем_НеСходится()
    {
        // Настоящий файл старой версии с его настоящей подписью нельзя выдать за новую версию.
        var (publicKey, signature) = Sign("WarehousePlanAutomation-1.19.0-win-x64.exe", Content);

        Assert.False(UpdateSignature.IsValid(Name, new MemoryStream(Content), signature, publicKey));
    }

    [Fact]
    public void ПодписьДругимКлючом_НеСходится()
    {
        var (_, signature) = Sign(Name, Content);
        var (otherKey, _) = Sign(Name, Content);

        Assert.False(UpdateSignature.IsValid(Name, new MemoryStream(Content), signature, otherKey));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("не base64 !!!")]
    [InlineData("AAAA")]
    public void МусорВместоПодписи_НеСходится(string? signature)
    {
        Assert.False(UpdateSignature.IsValid(Name, new MemoryStream(Content), signature));
    }

    [Fact]
    public void ВстроенныйКлюч_Читается()
    {
        using var key = System.Security.Cryptography.ECDsa.Create();
        key.ImportSubjectPublicKeyInfo(Convert.FromBase64String(UpdateSignature.PublicKey), out _);

        Assert.Equal(256, key.KeySize);
    }
}
