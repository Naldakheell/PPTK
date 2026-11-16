using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

/// <summary>The live slideshow check that replaced the cached snapshot for the update notice (2C-4 finding: the
/// snapshot still said "slideshow" for 30 s after the show ended). Real windows in the test process, no hook.</summary>
[SupportedOSPlatform("windows")]
public class PowerPointWindowsTests
{
    [Theory]
    [InlineData("screenClass", true)]
    [InlineData("PodiumParent", true)]
    [InlineData("PPTFrameClass", false)]
    [InlineData("NUIDialog", false)]
    [InlineData("#32770", false)]
    [InlineData("", false)]
    public void Slideshow_window_classes(string cls, bool slideshow) => Assert.Equal(slideshow, PowerPointWindows.IsSlideshowClass(cls));

    delegate IntPtr WndProc(IntPtr h, uint m, IntPtr w, IntPtr l);
    static readonly WndProc Proc = (h, m, w, l) => DefWindowProc(h, m, w, l);
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct WndClass { public uint style; public IntPtr lpfnWndProc; public int cbClsExtra, cbWndExtra; public IntPtr hInstance, hIcon, hCursor, hbrBackground; public string? lpszMenuName; public string lpszClassName; }
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern ushort RegisterClass(ref WndClass wc);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr DefWindowProc(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern IntPtr CreateWindowEx(int ex, string cls, string name, int style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr inst, IntPtr p);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] static extern bool DestroyWindow(IntPtr h);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern IntPtr GetModuleHandle(string? name);

    // A window of the given class in this process, far off screen, never activated (WS_EX_NOACTIVATE | TOOLWINDOW).
    static IntPtr Window(string cls)
    {
        var wc = new WndClass { lpfnWndProc = Marshal.GetFunctionPointerForDelegate(Proc), hInstance = GetModuleHandle(null), lpszClassName = cls };
        RegisterClass(ref wc);   // 1410 (already registered) is fine
        var h = CreateWindowEx(0x08000000 | 0x80, cls, "PPTK test window", unchecked((int)0x80000000), -32000, -32000, 10, 10, IntPtr.Zero, IntPtr.Zero, GetModuleHandle(null), IntPtr.Zero);
        Assert.NotEqual(IntPtr.Zero, h);
        return h;
    }

    [Fact]
    public void A_visible_slideshow_window_is_seen_live_and_its_end_is_seen_at_once()
    {
        using var ui = new UiThread();
        uint me = (uint)Environment.ProcessId;
        var show = ui.Run(() => Window("screenClass"));
        try
        {
            Assert.False(PowerPointWindows.HasVisibleSlideshow(me));                  // created hidden: no show yet
            ui.Run(() => ShowWindow(show, 4));                                        // SW_SHOWNOACTIVATE
            Assert.True(PowerPointWindows.HasVisibleSlideshow(me));                   // show running
            Assert.False(PowerPointWindows.HasVisibleSlideshow(me + 7919));           // another process's windows do not count
            Assert.False(PowerPointWindows.HasVisibleSlideshow(0));
            ui.Run(() => ShowWindow(show, 0));                                        // SW_HIDE: show ended
            Assert.False(PowerPointWindows.HasVisibleSlideshow(me));                  // no stale "slideshow"
            ui.Run(() => ShowWindow(show, 4));
            Assert.True(PowerPointWindows.HasVisibleSlideshow(me));
        }
        finally { ui.Run(() => DestroyWindow(show)); }
        Assert.False(PowerPointWindows.HasVisibleSlideshow(me));                      // window gone: no show
    }

    [Fact]
    public void Presenter_view_counts_and_an_ordinary_frame_does_not()
    {
        using var ui = new UiThread();
        uint me = (uint)Environment.ProcessId;
        var frame = ui.Run(() => { var h = Window("PPTFrameClass"); ShowWindow(h, 4); return h; });
        try
        {
            Assert.False(PowerPointWindows.HasVisibleSlideshow(me));                  // editing window only
            var podium = ui.Run(() => { var h = Window("PodiumParent"); ShowWindow(h, 4); return h; });
            try { Assert.True(PowerPointWindows.HasVisibleSlideshow(me)); }
            finally { ui.Run(() => DestroyWindow(podium)); }
            Assert.False(PowerPointWindows.HasVisibleSlideshow(me));
        }
        finally { ui.Run(() => DestroyWindow(frame)); }
    }

    [Fact]
    public void Nothing_to_announce_means_no_observation_at_all()
    {
        int observed = 0;
        var c = new UpdateNoticeController(() => { observed++; return new NoticeObservation(true, false, false, TimeSpan.FromMinutes(1)); },
            _ => true, _ => MarkResult.Marked, w => w(), _ => { });
        c.Tick();
        c.OnInspection(new PendingInspection(false, false, null, false, "no pending update")); c.Tick();
        c.OnInspection(new PendingInspection(true, false, "1.0.1", false, "bad hash")); c.Tick();
        c.OnInspection(new PendingInspection(true, true, "1.0.1", true, null)); c.Tick();
        Assert.Equal(0, observed);                                                     // no window enumeration when nothing is due
        c.OnInspection(new PendingInspection(true, true, "1.0.1", false, null)); c.Tick();
        Assert.Equal(1, observed);
    }
}
