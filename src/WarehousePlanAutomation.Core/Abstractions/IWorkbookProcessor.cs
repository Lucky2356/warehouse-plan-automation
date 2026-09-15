namespace WarehousePlanAutomation.Core.Abstractions;

/// <summary>
/// Счётчик, который показывается в карточке результата.
/// </summary>
/// <param name="NeedsAttention">
/// Значение требует действия человека. Такой счётчик окно подсвечивает отдельно:
/// остальные числа сообщают, что программа сделала, и смотреть на них не обязательно.
/// </param>
public sealed record ProcessingCounter(string Caption, int Value, bool NeedsAttention = false);

/// <summary>
/// Замечание, которое программа не стала решать за человека.
/// </summary>
/// <param name="Where">Место в книге - лист и ячейка или строка. null, если замечание общее.</param>
/// <param name="Sheet">
/// Название листа так, как оно записано в книге. Заполнено там, где программа знает,
/// куда смотреть: по нему окно открывает готовую книгу сразу на нужном месте.
/// </param>
/// <param name="Cell">
/// Адрес ячейки на этом листе («BD5»). Без него открывается просто лист.
/// </param>
public sealed record ProcessingWarning(
    string Message,
    string? Where = null,
    string? Sheet = null,
    string? Cell = null)
{
    /// <summary>Есть куда перейти: замечание привязано хотя бы к листу.</summary>
    public bool CanNavigate => !string.IsNullOrEmpty(Sheet);
}

/// <summary>Итог обработки книги.</summary>
public sealed record ProcessingResult(
    string OutputPath,
    IReadOnlyList<ProcessingCounter> Counters,
    IReadOnlyList<ProcessingWarning> Warnings)
{
    public static ProcessingResult Empty(string outputPath) =>
        new(outputPath, Array.Empty<ProcessingCounter>(), Array.Empty<ProcessingWarning>());
}

/// <summary>Этап обработки для отображения в интерфейсе.</summary>
public sealed record ProcessingStage(string Message, int Percent);

/// <summary>Вопрос, на который программа не может ответить сама.</summary>
/// <param name="Question">Что именно спрашивается, одной фразой.</param>
/// <param name="Explanation">Почему выбор нельзя сделать автоматически.</param>
/// <param name="Options">Варианты ответа в том виде, в каком они записаны в книге.</param>
public sealed record DecisionRequest(
    string Question,
    string Explanation,
    IReadOnlyList<string> Options);

/// <summary>
/// Запрос решения у человека посреди обработки. Обработка идёт на своём потоке,
/// поэтому реализация обязана сама переключиться на поток интерфейса.
/// Возврат null означает «человек не выбрал»: обработка продолжается без этого значения.
/// </summary>
public interface IDecisionPrompt
{
    string? Choose(DecisionRequest request);
}

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
