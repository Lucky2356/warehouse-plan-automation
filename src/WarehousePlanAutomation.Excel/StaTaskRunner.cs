namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Выполнение работы на выделенном STA-потоке. Microsoft Excel - однопоточный COM-сервер,
/// поэтому автоматизация выполняется вне потока интерфейса, но в подходящей апартаментной модели.
/// </summary>
internal static class StaTaskRunner
{
    public static Task<T> RunAsync<T>(Func<T> work, CancellationToken cancellationToken)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                completion.SetResult(work());
            }
            catch (OperationCanceledException)
            {
                completion.SetCanceled(cancellationToken);
            }
            catch (Exception ex)
            {
                completion.SetException(ex);
            }
        })
        {
            IsBackground = false,
            Name = "ExcelAutomation",
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        return completion.Task;
    }
}
