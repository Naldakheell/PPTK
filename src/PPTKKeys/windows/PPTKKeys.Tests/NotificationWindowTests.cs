using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

/// <summary>A thread with a message loop (like the helper's message thread) for the notice window tests.</summary>
[SupportedOSPlatform("windows")]
sealed class UiThread : IDisposable
{
    readonly ConcurrentQueue<Action> work = new();
    readonly Thread thread;
    volatile bool stop;

    public UiThread()
    {
        thread = new Thread(() =>
        {
            while (!stop)
            {
                while (work.TryDequeue(out var w)) w();
                while (Win.PeekMessage(out var msg, IntPtr.Zero, 0, 0, 1)) { Win.TranslateMessage(ref msg); Win.DispatchMessage(ref msg); }
                Thread.Sleep(5);
            }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public T Run<T>(Func<T> f)
    {
        T result = default!; Exception? error = null;
        using var done = new ManualResetEventSlim();
        work.Enqueue(() => { try { result = f(); } catch (Exception e) { error = e; } finally { done.Set(); } });
        Assert.True(done.Wait(10_000), "UI thread did not respond");
        if (error is not null) throw error;
        return result;
    }

    public void Dispose() { stop = true; thread.Join(5000); }
}

static class Win
{
    [StructLayout(LayoutKind.Sequential)] public struct Msg { public IntPtr hwnd; public uint message; public IntPtr wParam, lParam; public uint time; public int x, y; }
    [DllImport("user32.dll")] public static extern bool PeekMessage(out Msg m, IntPtr hwnd, uint min, uint max, uint remove);
    [DllImport("user32.dll")] public static extern bool TranslateMessage(ref Msg m);
    [DllImport("user32.dll")] public static extern IntPtr DispatchMessage(ref Msg m);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern IntPtr GetFocus();
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] public static extern IntPtr GetWindowLongPtr(IntPtr h, int index);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out NotificationWindow.Rect r);
    [DllImport("user32.dll")] public static extern bool DestroyWindow(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr CreateWindowEx(int ex, string cls, string name, int style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr inst, IntPtr p);
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc f, IntPtr l);
}

[SupportedOSPlatform("windows")]
public class NotificationWindowTests
{
    const int WsExTopmost = 0x8, WsExToolWindow = 0x80, WsExNoActivate = 0x08000000;
    const int WsPopup = unchecked((int)0x80000000), WsCaption = 0x00C00000, WsSysMenu = 0x00080000, WsChild = 0x40000000;

    static bool WaitUntil(Func<bool> condition, int ms)
    {
        var until = DateTime.UtcNow.AddMilliseconds(ms);
        while (DateTime.UtcNow < until) { if (condition()) return true; Thread.Sleep(20); }
        return condition();
    }

    [Fact]
    public void Notice_appears_without_taking_focus_or_the_foreground()
    {
        using var ui = new UiThread();
        var foregroundBefore = Win.GetForegroundWindow();
        var w = ui.Run(() => NotificationWindow.Show(UpdateNoticeText.Message, TimeSpan.FromSeconds(3), IntPtr.Zero));
        try
        {
            Assert.NotNull(w);                                         // non-null = it actually appeared
            Assert.True(Win.IsWindowVisible(w!.Handle));
            Thread.Sleep(150);
            var foregroundAfter = Win.GetForegroundWindow();
            Assert.NotEqual(w.Handle, foregroundAfter);
            Assert.Equal(foregroundBefore, foregroundAfter);            // whatever had the foreground keeps it
            Assert.NotEqual(w.Handle, ui.Run(Win.GetFocus));            // no keyboard focus, even on its own thread
        }
        finally { ui.Run(() => { w?.Close(); return 0; }); }
    }

    [Fact]
    public void Notice_is_a_non_activating_topmost_popup_with_text_and_no_buttons()
    {
        using var ui = new UiThread();
        var w = ui.Run(() => NotificationWindow.Show(UpdateNoticeText.Message, TimeSpan.FromSeconds(3), IntPtr.Zero))!;
        try
        {
            long ex = (long)Win.GetWindowLongPtr(w.Handle, -20), style = (long)Win.GetWindowLongPtr(w.Handle, -16);
            Assert.True((ex & WsExNoActivate) != 0 && (ex & WsExTopmost) != 0 && (ex & WsExToolWindow) != 0, $"exstyle {ex:X}");
            Assert.True((style & WsPopup) != 0);
            Assert.True((style & (WsCaption | WsSysMenu | WsChild)) == 0, $"style {style:X}");   // no title bar or system buttons
            int children = 0;
            Win.EnumChildWindows(w.Handle, (_, _) => { children++; return true; }, IntPtr.Zero);
            Assert.Equal(0, children);                                   // no button controls
            var sb = new StringBuilder(256);
            Win.GetWindowText(w.Handle, sb, sb.Capacity);
            Assert.Equal(UpdateNoticeText.Message, sb.ToString());       // also what a screen reader gets
            Assert.Equal(new IntPtr(3), ui.Run(() => Win.SendMessage(w.Handle, 0x0021, IntPtr.Zero, IntPtr.Zero)));   // WM_MOUSEACTIVATE -> MA_NOACTIVATE
        }
        finally { ui.Run(() => { w.Close(); return 0; }); }
    }

    [Fact]
    public void Notice_closes_itself_after_its_duration()
    {
        using var ui = new UiThread();
        NoticeCloseReason? reason = null;
        var w = ui.Run(() => { var x = NotificationWindow.Show(UpdateNoticeText.Message, TimeSpan.FromMilliseconds(400), IntPtr.Zero)!; x.Closed += r => reason = r; return x; });
        Assert.True(WaitUntil(() => reason is not null, 5000), "did not close");
        Assert.Equal(NoticeCloseReason.Timeout, reason);
        Assert.False(Win.IsWindow(w.Handle == IntPtr.Zero ? new IntPtr(1) : w.Handle) && w.IsOpen);
        Assert.False(w.IsOpen);
    }

    [Fact]
    public void A_click_closes_the_notice()
    {
        using var ui = new UiThread();
        NoticeCloseReason? reason = null;
        var w = ui.Run(() => { var x = NotificationWindow.Show(UpdateNoticeText.Message, TimeSpan.FromSeconds(10), IntPtr.Zero)!; x.Closed += r => reason = r; return x; });
        IntPtr h = w.Handle;
        Win.PostMessage(h, 0x0202, IntPtr.Zero, IntPtr.Zero);       // WM_LBUTTONUP (a message, not injected input)
        Assert.True(WaitUntil(() => reason is not null, 3000), "did not close on click");
        Assert.Equal(NoticeCloseReason.Clicked, reason);
        Assert.False(Win.IsWindow(h));
    }

    [Fact]
    public void Notice_sits_bottom_right_inside_the_PowerPoint_window_and_within_the_screen()
    {
        using var ui = new UiThread();
        var area = NotificationWindow.WorkArea(IntPtr.Zero);
        // A hidden stand-in for the PowerPoint frame, inside the work area.
        var frame = ui.Run(() => Win.CreateWindowEx(0, "STATIC", "PPTK test frame (hidden)", WsPopup, area.Left + 60, area.Top + 60, 900, 600, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero));
        try
        {
            Assert.True(Win.GetWindowRect(frame, out var f));
            var w = ui.Run(() => NotificationWindow.Show(UpdateNoticeText.Message, TimeSpan.FromSeconds(3), frame))!;
            try
            {
                Assert.True(Win.GetWindowRect(w.Handle, out var n));
                Assert.True(n.Left >= f.Left && n.Right <= f.Right && n.Top >= f.Top && n.Bottom <= f.Bottom, $"notice {n.Left},{n.Top}-{n.Right},{n.Bottom} frame {f.Left},{f.Top}-{f.Right},{f.Bottom}");
                Assert.True(n.Right > f.Left + (f.Right - f.Left) / 2 && n.Bottom > f.Top + (f.Bottom - f.Top) / 2);   // bottom-right quarter
            }
            finally { ui.Run(() => { w.Close(); return 0; }); }

            // A frame partly off the screen: the notice stays within the work area.
            var off = NotificationWindow.Place(IntPtr.Zero, area, 340, 92, 16);
            Assert.True(off.Left >= area.Left && off.Right <= area.Right && off.Top >= area.Top && off.Bottom <= area.Bottom);
        }
        finally { ui.Run(() => Win.DestroyWindow(frame)); }
    }

    [Fact]
    public void Several_notices_and_explicit_close_report_their_reasons()
    {
        using var ui = new UiThread();
        var reasons = new ConcurrentBag<NoticeCloseReason>();
        var a = ui.Run(() => { var x = NotificationWindow.Show(UpdateNoticeText.Message, TimeSpan.FromSeconds(10), IntPtr.Zero)!; x.Closed += reasons.Add; return x; });
        var b = ui.Run(() => { var x = NotificationWindow.Show(UpdateNoticeText.Message, TimeSpan.FromMilliseconds(300), IntPtr.Zero)!; x.Closed += reasons.Add; return x; });
        ui.Run(() => { a.Close(); return 0; });
        Assert.True(WaitUntil(() => reasons.Count == 2, 5000));
        Assert.Contains(NoticeCloseReason.Closed, reasons);
        Assert.Contains(NoticeCloseReason.Timeout, reasons);
        Assert.False(a.IsOpen || b.IsOpen);
    }
}
