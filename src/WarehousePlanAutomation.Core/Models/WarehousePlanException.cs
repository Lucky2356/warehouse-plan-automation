namespace WarehousePlanAutomation.Core.Models;

/// <summary>Ошибка бизнес-уровня с текстом, готовым для показа пользователю.</summary>
public class WarehousePlanException : Exception
{
    public WarehousePlanException(string message)
        : base(message)
    {
    }

    public WarehousePlanException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

/// <summary>Ошибка структуры книги: не найдены листы, заголовки или служебные строки.</summary>
public sealed class WorkbookValidationException : WarehousePlanException
{
    public WorkbookValidationException(IReadOnlyList<string> problems)
        : base(BuildMessage(problems))
    {
        Problems = problems;
    }

    public IReadOnlyList<string> Problems { get; }

    private static string BuildMessage(IReadOnlyList<string> problems)
    {
        if (problems.Count == 0)
        {
            return "Структура файла не соответствует ожидаемой.";
        }

        return "Структура файла не соответствует ожидаемой:" + Environment.NewLine +
               string.Join(Environment.NewLine, problems.Select(p => " - " + p));
    }
}
