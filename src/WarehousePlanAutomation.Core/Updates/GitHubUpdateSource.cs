using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using WarehousePlanAutomation.Core.Logging;

namespace WarehousePlanAutomation.Core.Updates;

/// <summary>Откуда программа узнаёт о новых выпусках.</summary>
public interface IUpdateSource
{
    /// <summary>Страница выпусков - её открывает кнопка «Открыть на GitHub».</summary>
    string ReleasesPage { get; }

    /// <summary>Последний выпуск или причина, по которой узнать не удалось.</summary>
    Task<UpdateCheck> CheckAsync(Version current, CancellationToken cancellationToken);

    /// <summary>Скачивает файл выпуска и возвращает путь к нему.</summary>
    Task<string> DownloadAsync(
        AppRelease release,
        IProgress<int>? progress,
        CancellationToken cancellationToken);
}

/// <summary>
/// Выпуски берутся из открытого репозитория на GitHub - того же, куда рабочий процесс
/// выкладывает готовый .exe. Ни ключей, ни входа в учётную запись не нужно: репозиторий
/// открытый, и запрос идёт обычным анонимным GET.
/// </summary>
public sealed class GitHubUpdateSource : IUpdateSource, IDisposable
{
    /// <summary>GitHub отказывает запросам без User-Agent, поэтому он обязателен.</summary>
    private const string UserAgent = "WarehousePlanAutomation";

    private static readonly TimeSpan CheckTimeout = TimeSpan.FromSeconds(20);

    private readonly IAppLogger _logger;
    private readonly string _owner;
    private readonly string _repository;
    private readonly HttpClient _client;

    public GitHubUpdateSource(IAppLogger logger, string owner, string repository)
    {
        _logger = logger;
        _owner = owner;
        _repository = repository;

        _client = new HttpClient(new HttpClientHandler
        {
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
        });

        _client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue(UserAgent, "1.0"));
        _client.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
    }

    public string ReleasesPage =>
        "https://github.com/" + _owner + "/" + _repository + "/releases";

    private string LatestReleaseUrl =>
        "https://api.github.com/repos/" + _owner + "/" + _repository + "/releases/latest";

    public async Task<UpdateCheck> CheckAsync(Version current, CancellationToken cancellationToken)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(CheckTimeout);

        HttpResponseMessage response;
        try
        {
            response = await _client.GetAsync(LatestReleaseUrl, timeout.Token).ConfigureAwait(false);
        }
        catch (HttpRequestException ex)
        {
            _logger.Warning("Не удалось запросить обновления.", ex);
            return UpdateCheck.Failed("Нет связи с GitHub. Проверьте подключение к интернету.");
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            _logger.Warning("Запрос обновлений не уложился в отведённое время.");
            return UpdateCheck.Failed("GitHub не ответил за 20 секунд. Попробуйте ещё раз.");
        }

        using (response)
        {
            if (response.StatusCode == HttpStatusCode.NotFound)
            {
                _logger.Warning("GitHub ответил 404 на запрос выпусков: " + LatestReleaseUrl);
                return UpdateCheck.Failed(
                    "GitHub не отдал список выпусков: репозиторий закрыт или выпусков ещё нет.");
            }

            if (!response.IsSuccessStatusCode)
            {
                _logger.Warning("GitHub ответил " + (int)response.StatusCode + " на запрос выпусков.");
                return UpdateCheck.Failed(
                    "GitHub ответил ошибкой " + (int)response.StatusCode + ". Попробуйте позже.");
            }

            var json = await response.Content.ReadAsStringAsync(timeout.Token).ConfigureAwait(false);
            var release = GitHubReleaseReader.Read(json);

            if (release is null)
            {
                _logger.Warning("Последний выпуск разобрать не удалось.");
                return UpdateCheck.Failed("У последнего выпуска не нашлось готового .exe.");
            }

            _logger.Information(
                "Последний выпуск: " + release.Tag + ", файл " + release.AssetName +
                " (" + release.AssetSize + " байт). Текущая версия: " + AppVersion.Display(current) + ".");

            return AppVersion.IsNewer(release.Version, current)
                ? new UpdateCheck(release, null)
                : UpdateCheck.None();
        }
    }

    public async Task<string> DownloadAsync(
        AppRelease release,
        IProgress<int>? progress,
        CancellationToken cancellationToken)
    {
        var directory = Path.Combine(Path.GetTempPath(), UserAgent, "update");
        Directory.CreateDirectory(directory);

        // Незавершённая закачка прошлого раза не должна остаться и тем более
        // не должна быть запущена: папка очищается перед каждой загрузкой.
        foreach (var stale in Directory.EnumerateFiles(directory))
        {
            try
            {
                File.Delete(stale);
            }
            catch (IOException ex)
            {
                _logger.Warning("Не удалось удалить прошлую загрузку: " + stale, ex);
            }
        }

        // Без подписи файл не качается вовсе: проверить его будет не по чему.
        if (string.IsNullOrEmpty(release.SignatureUrl))
        {
            _logger.Warning("У выпуска " + release.Tag + " нет файла подписи " + release.AssetName + UpdateSignature.Extension + ".");
            throw new UpdateSignatureException(
                "У выпуска " + release.Tag + " нет подписи. Если он вышел только что, подпись появится в течение часа.");
        }

        var path = Path.Combine(directory, release.AssetName);
        var signaturePath = path + UpdateSignature.Extension;
        var signature = await DownloadSignatureAsync(release.SignatureUrl, cancellationToken).ConfigureAwait(false);
        await File.WriteAllTextAsync(signaturePath, signature, cancellationToken).ConfigureAwait(false);

        _logger.Information("Загрузка обновления в " + path);

        using var response = await _client
            .GetAsync(release.AssetUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
            .ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        var total = response.Content.Headers.ContentLength ?? release.AssetSize;

        await using (var source = await response.Content
                         .ReadAsStreamAsync(cancellationToken).ConfigureAwait(false))
        await using (var target = new FileStream(
                         path, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16, true))
        {
            var buffer = new byte[1 << 16];
            long done = 0;
            var reported = -1;

            while (true)
            {
                var read = await source.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
                if (read == 0)
                {
                    break;
                }

                await target.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
                done += read;

                if (total <= 0)
                {
                    continue;
                }

                var percent = (int)(done * 100 / total);
                if (percent != reported)
                {
                    reported = percent;
                    progress?.Report(percent);
                }
            }
        }

        var size = new FileInfo(path).Length;
        if (total > 0 && size != total)
        {
            File.Delete(path);
            throw new IOException(
                "Файл обновления скачан не полностью: " + size + " байт вместо " + total + ".");
        }

        try
        {
            UpdateSignature.EnsureValid(path, release.AssetName, signature);
        }
        catch (UpdateSignatureException)
        {
            _logger.Error("Подпись обновления " + release.AssetName + " не сошлась - файл удалён и не будет установлен.");
            File.Delete(path);
            File.Delete(signaturePath);
            throw;
        }

        _logger.Information("Обновление скачано: " + size + " байт, подпись сошлась.");
        return path;
    }

    /// <summary>Подпись - несколько десятков байт; больше - это уже не подпись.</summary>
    private async Task<string> DownloadSignatureAsync(string url, CancellationToken cancellationToken)
    {
        using var response = await _client.GetAsync(url, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new UpdateSignatureException(
                "Подпись выпуска не скачалась (GitHub ответил " + (int)response.StatusCode + ").");
        }

        var bytes = await response.Content.ReadAsByteArrayAsync(cancellationToken).ConfigureAwait(false);
        if (bytes.Length > 4096)
        {
            throw new UpdateSignatureException("Файл подписи выпуска подозрительно большой - обновление не ставится.");
        }

        return System.Text.Encoding.ASCII.GetString(bytes);
    }

    public void Dispose() => _client.Dispose();
}
