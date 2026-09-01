using System.Runtime.InteropServices;

namespace WarehousePlanAutomation.Excel;

/// <summary>Освобождение COM-объектов Excel.</summary>
internal static class ComUtils
{
    /// <summary>
    /// Сбалансированное освобождение: одна обёртка - один вызов.
    /// Excel может вернуть один и тот же COM-объект несколько раз (например, коллекцию листов),
    /// и тогда среда выполнения отдаёт одну и ту же обёртку с увеличенным счётчиком ссылок.
    /// Поэтому в обычном коде используется ReleaseComObject, а не FinalReleaseComObject:
    /// иначе чужая, ещё используемая ссылка была бы разорвана.
    /// </summary>
    public static void Release(object? comObject)
    {
        if (comObject is null)
        {
            return;
        }

        try
        {
            if (Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch (ArgumentException)
        {
            // Объект уже не является COM-обёрткой - освобождать нечего.
        }
        catch (InvalidComObjectException)
        {
            // Обёртка уже разъединена с COM-объектом.
        }
    }

    /// <summary>
    /// Полное освобождение обёртки. Применяется только к объектам, которые точно больше
    /// нигде не используются (приложение Excel после Quit), чтобы не осталось живых ссылок,
    /// удерживающих процесс EXCEL.EXE.
    /// </summary>
    public static void FinalRelease(object? comObject)
    {
        if (comObject is null)
        {
            return;
        }

        try
        {
            if (Marshal.IsComObject(comObject))
            {
                Marshal.FinalReleaseComObject(comObject);
            }
        }
        catch (ArgumentException)
        {
            // Объект уже не является COM-обёрткой.
        }
        catch (InvalidComObjectException)
        {
            // Обёртка уже разъединена с COM-объектом.
        }
    }

    /// <summary>Принудительный сбор мусора, чтобы не осталось живых обёрток RCW.</summary>
    public static void CollectGarbage()
    {
        for (var i = 0; i < 2; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}

/// <summary>
/// Область видимости COM-объектов: все взятые обёртки освобождаются в обратном порядке.
/// Позволяет не оставлять неосвобождённых временных объектов в цепочках вызовов.
/// </summary>
internal sealed class ComScope : IDisposable
{
    private readonly List<object> _items = new();

    public dynamic Track(object? value)
    {
        if (value is not null)
        {
            _items.Add(value);
        }

        return value!;
    }

    public void Dispose()
    {
        for (var i = _items.Count - 1; i >= 0; i--)
        {
            ComUtils.Release(_items[i]);
        }

        _items.Clear();
    }
}
