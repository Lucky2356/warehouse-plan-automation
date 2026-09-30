using System.Runtime.InteropServices;

namespace WarehousePlanAutomation.Excel;

/// <summary>
/// Фильтр сообщений COM для STA-потока автоматизации. Если Excel занят (запускается, загружает
/// надстройки, показывает окно) и отклоняет вызов (RPC_E_CALL_REJECTED, 0x80010001),
/// вызов повторяется, а не завершается ошибкой.
/// </summary>
[ComVisible(true)]
internal sealed class ComMessageFilter : IOleMessageFilter, IDisposable
{
    private const int ServerCallIsHandled = 0;
    private const int ServerCallRetryLater = 2;
    private const int PendingMessageWaitDefProcess = 2;
    private const int RetryImmediately = -1;
    private const int CancelCall = -1;
    private const int RetryDelayMilliseconds = 200;
    private const int MaxWaitMilliseconds = 120_000;

    private readonly IOleMessageFilter? _previous;
    private bool _registered;

    private ComMessageFilter()
    {
        if (CoRegisterMessageFilter(this, out _previous) == 0)
        {
            _registered = true;
        }
    }

    /// <summary>Регистрирует фильтр для текущего STA-потока. Освобождение снимает регистрацию.</summary>
    public static ComMessageFilter Register() => new();

    public void Dispose()
    {
        if (!_registered)
        {
            return;
        }

        _registered = false;
        _ = CoRegisterMessageFilter(_previous, out _);
    }

    int IOleMessageFilter.HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr interfaceInfo) =>
        ServerCallIsHandled;

    int IOleMessageFilter.RetryRejectedCall(IntPtr callee, int tickCount, int rejectType)
    {
        // Сервер занят: повторяем через паузу, но не бесконечно.
        if (rejectType == ServerCallRetryLater && tickCount < MaxWaitMilliseconds)
        {
            return RetryDelayMilliseconds;
        }

        return CancelCall;
    }

    int IOleMessageFilter.MessagePending(IntPtr callee, int tickCount, int pendingType) =>
        PendingMessageWaitDefProcess;

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);
}

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr interfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr callee, int tickCount, int rejectType);

    [PreserveSig]
    int MessagePending(IntPtr callee, int tickCount, int pendingType);
}
