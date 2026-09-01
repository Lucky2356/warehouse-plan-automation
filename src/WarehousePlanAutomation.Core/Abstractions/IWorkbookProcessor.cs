namespace WarehousePlanAutomation.Core.Abstractions;

/// <summary>Итог обработки книги.</summary>
public sealed record ProcessingResult(
    string OutputPath,
    int ProcessedOrderRows,
    int RemainingOrderRows,
    int NewPlanOrders,
    int DeletedPlaceholderRows);

/// <summary>Этап обработки для отображения в интерфейсе.</summary>
public sealed record ProcessingStage(string Message, int Percent);

/// <summary>
/// Обработчик книги. Реализация работает через COM-автоматизацию Microsoft Excel,
/// но интерфейс от неё не зависит, поэтому бизнес-логика тестируется без установленного Office.
/// </summary>
public interface IWorkbookProcessor
{
    Task<ProcessingResult> ProcessAsync(
        string sourceFilePath,
        IProgress<ProcessingStage>? progress,
        CancellationToken cancellationToken);
}
