using System.Text;

namespace PPTKKeys;

/// <summary>
/// Foreground and focus checks that are safe inside a low-level keyboard hook.
/// Ribbon, menu, and dialog vetoes are Win32 class and GUI-thread state, read on the key.
/// </summary>
internal static class FocusProbe
{
    const uint ProcessQueryLimitedInformation = 0x1000;

    static IntPtr _cachedForeground;
    static bool _cachedIsPowerPoint;
    static uint _cachedPid;

    public readonly record struct Probe(bool IsPowerPoint, FocusKind Focus, uint ProcessId, string Detail);

    public static Probe Inspect()
    {
        IntPtr foreground = NativeMethods.GetForegroundWindow();
        if (foreground == IntPtr.Zero)
            return new Probe(false, FocusKind.Outside, 0, "no foreground");

        if (foreground != _cachedForeground)
        {
            NativeMethods.GetWindowThreadProcessId(foreground, out uint pid);
            _cachedPid = pid;
            _cachedIsPowerPoint = IsPowerPointProcess(pid);
            _cachedForeground = foreground;
        }

        if (!_cachedIsPowerPoint)
            return new Probe(false, FocusKind.Outside, _cachedPid, "not PowerPoint");

        string frame = ClassName(foreground);
        if (!frame.Equals("PPTFrameClass", StringComparison.Ordinal))
            return new Probe(true, FocusKind.NotFrame, _cachedPid, "class=" + frame);

        uint threadId = NativeMethods.GetWindowThreadProcessId(foreground, out _);
        var info = new NativeMethods.GuiThreadInfo
        {
            cbSize = System.Runtime.InteropServices.Marshal.SizeOf<NativeMethods.GuiThreadInfo>(),
        };
        if (!NativeMethods.GetGUIThreadInfo(threadId, ref info))
            return new Probe(true, FocusKind.Uncertain, _cachedPid, "gui thread unavailable");

        if ((info.flags & (NativeMethods.GuiInMenuMode | NativeMethods.GuiSystemMenuMode | NativeMethods.GuiPopupMenuMode)) != 0)
            return new Probe(true, FocusKind.MenuOrDialog, _cachedPid, "menu mode");

        if (info.hwndFocus == IntPtr.Zero)
            return new Probe(true, FocusKind.Uncertain, _cachedPid, "no focus window");

        bool sawPane = false;
        var chain = new StringBuilder();
        IntPtr hwnd = info.hwndFocus;
        for (int depth = 0; depth < 10 && hwnd != IntPtr.Zero; depth++)
        {
            string name = ClassName(hwnd);
            if (chain.Length > 0) chain.Append('/');
            chain.Append(name);
            if (IsBlockedClass(name))
                return new Probe(true, FocusKind.RibbonOrEdit, _cachedPid, "class=" + name);
            if (name is "paneClassDC" or "mdiClass") sawPane = true;
            if (hwnd == foreground) break;
            hwnd = NativeMethods.GetParent(hwnd);
        }

        if (!sawPane)
            return new Probe(true, FocusKind.Uncertain, _cachedPid, "chain=" + chain);

        return new Probe(true, FocusKind.Pane, _cachedPid, "chain=" + chain);
    }

    public static void Write(TextWriter writer)
    {
        Probe probe = Inspect();
        writer.WriteLine($"foregroundPowerPoint={probe.IsPowerPoint} pid={probe.ProcessId}");
        writer.WriteLine($"focus={probe.Focus} {probe.Detail}");
    }

    static bool IsPowerPointProcess(uint pid)
    {
        if (pid == 0) return false;
        IntPtr process = NativeMethods.OpenProcess(ProcessQueryLimitedInformation, false, pid);
        if (process == IntPtr.Zero) return false;
        try
        {
            int size = 1024;
            var buffer = new StringBuilder(size);
            if (!NativeMethods.QueryFullProcessImageName(process, 0, buffer, ref size)) return false;
            return Path.GetFileName(buffer.ToString()).Equals("POWERPNT.EXE", StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            NativeMethods.CloseHandle(process);
        }
    }

    static string ClassName(IntPtr hwnd)
    {
        var buffer = new StringBuilder(256);
        int length = NativeMethods.GetClassName(hwnd, buffer, buffer.Capacity);
        return length > 0 ? buffer.ToString() : "";
    }

    static bool IsBlockedClass(string name) => name switch
    {
        "Edit"
            or "RichEdit20A"
            or "RichEdit20W"
            or "RichEdit50W"
            or "RichEditD2DPT"
            or "RICHEDIT"
            or "RICHEDIT60W"
            or "NetUIHWND"
            or "NetUICtrlNotifySink"
            or "NUIPane"
            or "NUIDialog"
            or "MsoCommandBar"
            or "MsoCommandBarDock"
            or "MsoCommandBarPopup"
            or "#32770" => true,
        _ => name.StartsWith("bosa_sdm", StringComparison.Ordinal),
    };
}
