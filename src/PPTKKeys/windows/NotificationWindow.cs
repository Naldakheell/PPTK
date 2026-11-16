using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace PPTKKeys;

// Phase 2C-2: the update notice window. A small popup that never takes focus: WS_EX_NOACTIVATE, shown with
// SW_SHOWNOACTIVATE, and WM_MOUSEACTIVATE answered with MA_NOACTIVATE, so typing, dialogs and slideshows keep
// the keyboard. Text and a close mark only (no buttons); a click anywhere closes it, otherwise it closes itself.
// Must be created on a thread that pumps messages (the helper's message thread). Not connected yet (2C-3).

public enum NoticeCloseReason { Timeout, Clicked, Closed }

[SupportedOSPlatform("windows")]
public sealed class NotificationWindow
{
    const string ClassName = "PPTKKeys.UpdateNotice";
    const int WsPopup = unchecked((int)0x80000000), WsExTopmost = 0x8, WsExToolWindow = 0x80, WsExNoActivate = 0x08000000;
    const int SwShowNoActivate = 4, WmDestroy = 0x0002, WmPaint = 0x000F, WmTimer = 0x0113, WmLButtonUp = 0x0202, WmMouseActivate = 0x0021, WmNcHitTest = 0x0084;
    const int MaNoActivate = 3, HtClient = 1, CloseTimerId = 1;
    const uint SwpNoActivate = 0x10, SwpShowWindow = 0x40;
    static readonly IntPtr HwndTopmost = new(-1);
    static readonly WndProc Proc = OnMessage;   // kept alive for the window class
    static readonly Dictionary<IntPtr, NotificationWindow> Open = [];
    static bool registered;

    readonly string text;
    public IntPtr Handle { get; private set; }
    public NoticeCloseReason? ClosedBecause { get; private set; }
    public bool IsOpen => Handle != IntPtr.Zero && IsWindow(Handle);
    public event Action<NoticeCloseReason>? Closed;

    NotificationWindow(string text) { this.text = text; }

    /// <summary>Creates and shows the notice without activating it, bottom-right inside the anchor window (the
    /// PowerPoint frame) or the work area. Returns null when the window could not be created or is not visible:
    /// only a non-null result means the notice actually appeared.</summary>
    public static NotificationWindow? Show(string text, TimeSpan duration, IntPtr anchor)
    {
        if (!registered)
        {
            var wc = new WndClass { lpfnWndProc = Marshal.GetFunctionPointerForDelegate(Proc), hInstance = GetModuleHandle(null), lpszClassName = ClassName, hCursor = LoadCursor(IntPtr.Zero, 32512) };
            if (RegisterClass(ref wc) == 0 && Marshal.GetLastWin32Error() != 1410) return null;   // 1410: already registered
            registered = true;
        }
        var w = new NotificationWindow(text);
        int dpi = anchor != IntPtr.Zero ? (int)GetDpiForWindow(anchor) : 96;
        if (dpi <= 0) dpi = 96;
        int width = Scale(340, dpi), height = Scale(92, dpi), margin = Scale(16, dpi);
        var area = WorkArea(anchor);
        var r = Place(anchor, area, width, height, margin);
        IntPtr hwnd = CreateWindowEx(WsExNoActivate | WsExTopmost | WsExToolWindow, ClassName, text, WsPopup, r.Left, r.Top, width, height, IntPtr.Zero, IntPtr.Zero, GetModuleHandle(null), IntPtr.Zero);
        if (hwnd == IntPtr.Zero) return null;
        w.Handle = hwnd;
        Open[hwnd] = w;
        ShowWindow(hwnd, SwShowNoActivate);
        SetWindowPos(hwnd, HwndTopmost, 0, 0, 0, 0, 0x1 | 0x2 | SwpNoActivate | SwpShowWindow);   // SWP_NOSIZE | SWP_NOMOVE
        if (!IsWindowVisible(hwnd)) { w.Close(); return null; }
        SetTimer(hwnd, new IntPtr(CloseTimerId), (uint)Math.Max(1, duration.TotalMilliseconds), IntPtr.Zero);
        return w;
    }

    public void Close() => Close(NoticeCloseReason.Closed);

    void Close(NoticeCloseReason reason)
    {
        if (Handle == IntPtr.Zero) return;
        ClosedBecause ??= reason;
        if (IsWindow(Handle)) DestroyWindow(Handle);
    }

    static int Scale(int px, int dpi) => (int)Math.Round(px * dpi / 96.0);

    /// <summary>Bottom-right inside the anchor window, kept within the work area of its monitor.</summary>
    internal static Rect Place(IntPtr anchor, Rect area, int width, int height, int margin)
    {
        Rect within = area;
        if (anchor != IntPtr.Zero && GetWindowRect(anchor, out var a) && a.Right > a.Left && a.Bottom > a.Top) within = Intersect(a, area);
        int left = Math.Max(area.Left, within.Right - width - margin), top = Math.Max(area.Top, within.Bottom - height - margin);
        left = Math.Min(left, area.Right - width); top = Math.Min(top, area.Bottom - height);
        return new Rect { Left = left, Top = top, Right = left + width, Bottom = top + height };
    }

    static Rect Intersect(Rect a, Rect b)
    {
        var r = new Rect { Left = Math.Max(a.Left, b.Left), Top = Math.Max(a.Top, b.Top), Right = Math.Min(a.Right, b.Right), Bottom = Math.Min(a.Bottom, b.Bottom) };
        return r.Right > r.Left && r.Bottom > r.Top ? r : b;
    }

    internal static Rect WorkArea(IntPtr anchor)
    {
        IntPtr monitor = MonitorFromWindow(anchor, 2 /* MONITOR_DEFAULTTONEAREST */);
        var info = new MonitorInfo { cbSize = Marshal.SizeOf<MonitorInfo>() };
        return GetMonitorInfo(monitor, ref info) ? info.rcWork : new Rect { Left = 0, Top = 0, Right = 1280, Bottom = 720 };
    }

    static IntPtr OnMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        Open.TryGetValue(hwnd, out var w);
        switch (msg)
        {
            case WmMouseActivate: return new IntPtr(MaNoActivate);   // a click never activates it
            case WmNcHitTest: return new IntPtr(HtClient);
            case WmLButtonUp: w?.Close(NoticeCloseReason.Clicked); return IntPtr.Zero;
            case WmTimer when wParam == new IntPtr(CloseTimerId): w?.Close(NoticeCloseReason.Timeout); return IntPtr.Zero;
            case WmPaint: if (w is not null) w.Paint(hwnd); return IntPtr.Zero;
            case WmDestroy:
                KillTimer(hwnd, new IntPtr(CloseTimerId));
                if (w is not null)
                {
                    Open.Remove(hwnd);
                    w.Handle = IntPtr.Zero;
                    w.ClosedBecause ??= NoticeCloseReason.Closed;
                    w.Closed?.Invoke(w.ClosedBecause.Value);
                }
                return IntPtr.Zero;
        }
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    // Dark panel, white text, a close mark top-right (decoration only: a click anywhere closes).
    void Paint(IntPtr hwnd)
    {
        IntPtr dc = BeginPaint(hwnd, out var ps);
        try
        {
            GetClientRect(hwnd, out var rc);
            int dpi = (int)GetDpiForWindow(hwnd); if (dpi <= 0) dpi = 96;
            IntPtr back = CreateSolidBrush(0x00302A26), edge = CreateSolidBrush(0x00D77800);   // BGR
            FillRect(dc, ref rc, back);
            var bar = new Rect { Left = 0, Top = 0, Right = Scale(4, dpi), Bottom = rc.Bottom };
            FillRect(dc, ref bar, edge);
            DeleteObject(back); DeleteObject(edge);
            IntPtr font = CreateFont(-Scale(14, dpi), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, "Segoe UI");
            IntPtr old = SelectObject(dc, font);
            SetBkMode(dc, 1);   // transparent
            SetTextColor(dc, 0x00FFFFFF);
            var textRect = new Rect { Left = Scale(18, dpi), Top = Scale(12, dpi), Right = rc.Right - Scale(28, dpi), Bottom = rc.Bottom - Scale(8, dpi) };
            DrawText(dc, text.Replace("\n", "\r\n"), -1, ref textRect, 0x10 /* DT_WORDBREAK */);
            SetTextColor(dc, 0x00BBBBBB);
            var closeRect = new Rect { Left = rc.Right - Scale(24, dpi), Top = Scale(6, dpi), Right = rc.Right - Scale(6, dpi), Bottom = Scale(24, dpi) };
            DrawText(dc, "✕", -1, ref closeRect, 0x1 | 0x4 | 0x20);   // centred, single line
            SelectObject(dc, old); DeleteObject(font);
        }
        finally { EndPaint(hwnd, ref ps); }
    }

    // ------------------------------------------------------------ Win32 (user32/gdi32), local to this window
    delegate IntPtr WndProc(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct WndClass { public uint style; public IntPtr lpfnWndProc; public int cbClsExtra, cbWndExtra; public IntPtr hInstance, hIcon, hCursor, hbrBackground; public string? lpszMenuName; public string lpszClassName; }
    [StructLayout(LayoutKind.Sequential)] struct MonitorInfo { public int cbSize; public Rect rcMonitor, rcWork; public uint dwFlags; }
    [StructLayout(LayoutKind.Sequential)] struct PaintStruct { public IntPtr hdc; public bool fErase; public Rect rcPaint; public bool fRestore, fIncUpdate; [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)] public byte[] rgbReserved; }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern ushort RegisterClass(ref WndClass wc);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern IntPtr CreateWindowEx(int exStyle, string cls, string name, int style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr inst, IntPtr param);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool DestroyWindow(IntPtr hwnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr DefWindowProc(IntPtr hwnd, uint msg, IntPtr w, IntPtr l);   // W: the class is Unicode
    [DllImport("user32.dll")] static extern IntPtr SetTimer(IntPtr hwnd, IntPtr id, uint ms, IntPtr proc);
    [DllImport("user32.dll")] static extern bool KillTimer(IntPtr hwnd, IntPtr id);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out Rect r);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd, out Rect r);
    [DllImport("user32.dll")] static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);
    [DllImport("user32.dll")] static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr LoadCursor(IntPtr inst, int id);
    [DllImport("user32.dll")] static extern IntPtr BeginPaint(IntPtr hwnd, out PaintStruct ps);
    [DllImport("user32.dll")] static extern bool EndPaint(IntPtr hwnd, ref PaintStruct ps);
    [DllImport("user32.dll")] static extern int FillRect(IntPtr dc, ref Rect r, IntPtr brush);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int DrawText(IntPtr dc, string text, int count, ref Rect r, uint format);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern IntPtr GetModuleHandle(string? name);
    [DllImport("gdi32.dll")] static extern IntPtr CreateSolidBrush(uint color);
    [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr obj);
    [DllImport("gdi32.dll")] static extern IntPtr SelectObject(IntPtr dc, IntPtr obj);
    [DllImport("gdi32.dll")] static extern int SetBkMode(IntPtr dc, int mode);
    [DllImport("gdi32.dll")] static extern uint SetTextColor(IntPtr dc, uint color);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] static extern IntPtr CreateFont(int h, int w, int esc, int orient, int weight, uint italic, uint underline, uint strike, uint charset, uint outPrecision, uint clipPrecision, uint quality, uint pitch, string face);
}
