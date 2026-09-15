using WarehousePlanAutomation.Core.Logging;
using Xunit;

namespace WarehousePlanAutomation.Tests;

public class LogCleanupTests
{
    [Theory]
    [InlineData("2026-09-15 19:59", "2026-09-14 20:00")]
    [InlineData("2026-09-15 20:00", "2026-09-15 20:00")]
    [InlineData("2026-09-15 23:30", "2026-09-15 20:00")]
    [InlineData("2026-09-16 08:00", "2026-09-15 20:00")]
    public void ПоследнийВечернийСрок(string now, string boundary)
    {
        Assert.Equal(DateTime.Parse(boundary), LogCleanup.LastBoundary(DateTime.Parse(now)));
    }

    [Fact]
    public void СледующийСрок_ЧерезВечер()
    {
        Assert.Equal(new DateTime(2026, 9, 15, 20, 0, 0), LogCleanup.NextBoundary(new DateTime(2026, 9, 15, 10, 0, 0)));
        Assert.Equal(new DateTime(2026, 9, 16, 20, 0, 0), LogCleanup.NextBoundary(new DateTime(2026, 9, 15, 20, 5, 0)));
    }

    [Fact]
    public void Очистка_УдаляетЗаписанноеДоВечернегоСрока()
    {
        var directory = Path.Combine(Path.GetTempPath(), "log-cleanup-" + Guid.NewGuid().ToString("N"));
        try
        {
            var logger = new FileAppLogger(directory);
            var morning = Path.Combine(directory, "warehouse-plan-2026-09-15.log");
            var evening = Path.Combine(directory, "warehouse-plan-2026-09-16.log");
            var other = Path.Combine(directory, "чужой файл.txt");
            File.WriteAllText(morning, "утро");
            File.WriteAllText(evening, "вечер");
            File.WriteAllText(other, "не журнал");
            File.SetLastWriteTime(morning, new DateTime(2026, 9, 15, 14, 0, 0));
            File.SetLastWriteTime(evening, new DateTime(2026, 9, 15, 21, 0, 0));

            var removed = logger.RemoveExpired(new DateTime(2026, 9, 15, 22, 0, 0));

            Assert.Equal(1, removed);
            Assert.False(File.Exists(morning));
            Assert.True(File.Exists(evening));
            Assert.True(File.Exists(other));

            Assert.Equal(1, logger.RemoveExpired(new DateTime(2026, 9, 16, 20, 1, 0)));
            Assert.False(File.Exists(evening));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
