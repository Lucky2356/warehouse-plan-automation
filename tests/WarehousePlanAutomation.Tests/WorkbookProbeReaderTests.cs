using System.IO.Compression;
using System.Text;
using WarehousePlanAutomation.Core.Sheets;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class WorkbookProbeReaderTests : IDisposable
{
    private readonly string _folder =
        Path.Combine(Path.GetTempPath(), "wpa-probe-" + Guid.NewGuid().ToString("N"));

    public WorkbookProbeReaderTests() => Directory.CreateDirectory(_folder);

    public void Dispose()
    {
        try
        {
            Directory.Delete(_folder, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    /// <summary>Книга .xlsx - это zip; здесь собирается только та его часть, которую читает осмотр.</summary>
    private string MakeWorkbook(string sheets, string? externalTarget = null)
    {
        var path = Path.Combine(_folder, "книга.xlsx");

        using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            Write(archive, "xl/workbook.xml",
                """<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets>""" +
                sheets + "</sheets></workbook>");

            if (externalTarget is not null)
            {
                Write(archive, "xl/externalLinks/_rels/externalLink1.xml.rels",
                    """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target=""" +
                    "\"" + externalTarget + "\" TargetMode=\"External\"/></Relationships>");
            }
        }

        return path;
    }

    private static void Write(ZipArchive archive, string name, string content)
    {
        using var stream = archive.CreateEntry(name).Open();
        var bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }

    [Fact]
    public void ЧитаетНазванияЛистовВПорядкеКниги()
    {
        var path = MakeWorkbook(
            """<sheet name="прайс" sheetId="1" r:id="rId1" xmlns:r="x"/><sheet name="Цены" sheetId="2"/>""");

        var probe = WorkbookProbeReader.Read(path);

        Assert.NotNull(probe);
        Assert.Equal(new[] { "прайс", "Цены" }, probe!.Sheets);
    }

    [Fact]
    public void ЧитаетАдресСвязанногоФайла()
    {
        // Так Excel записывает ссылку на книгу в сетевой папке.
        var path = MakeWorkbook(
            """<sheet name="Цены" sheetId="1"/>""",
            "file:///\\\\FileServer\\%D1%81%D0%BA%D0%BB%D0%B0%D0%B4$\\FW26-27\\[стенки.xlsx]");

        var probe = WorkbookProbeReader.Read(path);

        Assert.NotNull(probe);
        Assert.Equal(@"\\FileServer\склад$\FW26-27\стенки.xlsx", Assert.Single(probe!.ExternalFiles));
    }

    [Fact]
    public void БезВнешнихСсылок_СписокПустой()
    {
        var probe = WorkbookProbeReader.Read(MakeWorkbook("""<sheet name="Цены" sheetId="1"/>"""));

        Assert.NotNull(probe);
        Assert.Empty(probe!.ExternalFiles);
    }

    [Fact]
    public void НеZip_ОсмотрНеДелается()
    {
        // Старый двоичный .xls так не прочитать - и это не ошибка, просто нет осмотра.
        var path = Path.Combine(_folder, "старая.xls");
        File.WriteAllText(path, "не zip");

        Assert.Null(WorkbookProbeReader.Read(path));
    }

    [Fact]
    public void ФайлаНет_ОсмотрНеДелается()
    {
        Assert.Null(WorkbookProbeReader.Read(Path.Combine(_folder, "нет.xlsx")));
    }
}
