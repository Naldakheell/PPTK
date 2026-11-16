using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;

namespace PPTKKeys;

/// <summary>Live, read-only checks of PowerPoint's top-level windows. Used by the update notice instead of the
/// cached PowerPoint snapshot, which is refreshed only on PowerPoint events and could still say "slideshow" after
/// the slideshow had ended (seen in the Phase 2C-4 real test).</summary>
[SupportedOSPlatform("windows")]
public static class PowerPointWindows
{
    /// <summary>Window classes of a running slideshow: the show itself and the presenter view.</summary>
    public static readonly IReadOnlySet<string> SlideshowClasses = new HashSet<string>(StringComparer.Ordinal) { "screenClass", "PodiumParent" };

    public static bool IsSlideshowClass(string windowClass) => SlideshowClasses.Contains(windowClass);

    /// <summary>True while the process has a visible slideshow or presenter-view window (on any monitor).</summary>
    public static bool HasVisibleSlideshow(uint processId)
    {
        if (processId == 0) return false;
        bool found = false;
        var name = new StringBuilder(64);
        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out uint pid);
            if (pid != processId || !IsWindowVisible(hwnd)) return true;
            name.Clear();
            if (GetClassName(hwnd, name, name.Capacity) > 0 && IsSlideshowClass(name.ToString())) { found = true; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder name, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
}
