using System.Runtime.InteropServices;

namespace PPTKKeys;

/// <summary>
/// WH_KEYBOARD_LL helper for the nine PPTK insert tokens.
/// The hook reads Win32 focus and the cached snapshot only. COM runs later, on this same
/// STA, when a posted message is dispatched. Any hook failure passes the key through.
/// </summary>
internal sealed class HelperApp
{
    const int QueueLimit = 128;

    readonly PowerPointSession _session = new();
    readonly StringBuilderPending _pending = new();
    readonly NativeMethods.LowLevelKeyboardProc _hookProc;
    readonly NativeMethods.WndProc _windowProc;
    readonly NativeMethods.WinEventProc _winEventProc;

    IntPtr _window;
    IntPtr _hook;
    IntPtr _foregroundHook;
    IntPtr _focusHook;
    IntPtr _lastFocus;
    uint _focusHookPid;
    int _heldVirtualKey;
    bool _capture;
    bool _needsRefresh;
    string _lastPass = "";

    internal HelperApp()
    {
        _hookProc = OnHook;
        _windowProc = OnWindow;
        _winEventProc = OnWinEvent;
    }

    public int Run()
    {
        _window = CreateMessageWindow();
        MessageWindow.Handle = _window;
        if (_window == IntPtr.Zero)
        {
            Console.Error.WriteLine($"The helper window was not created ({Marshal.GetLastWin32Error()}). No keys are intercepted.");
            return 1;
        }

        _hook = NativeMethods.SetWindowsHookEx(NativeMethods.WhKeyboardLl, _hookProc, IntPtr.Zero, 0);
        if (_hook == IntPtr.Zero)
        {
            Console.Error.WriteLine($"Keyboard hook failed ({Marshal.GetLastWin32Error()}). No keys are intercepted.");
            NativeMethods.DestroyWindow(_window);
            return 1;
        }

        _foregroundHook = NativeMethods.SetWinEventHook(
            NativeMethods.EventSystemForeground,
            NativeMethods.EventSystemForeground,
            IntPtr.Zero,
            _winEventProc,
            0,
            0,
            NativeMethods.WineventOutOfContext | NativeMethods.WineventSkipOwnProcess);
        NativeMethods.SetTimer(_window, new IntPtr(1), 1000, IntPtr.Zero);
        NativeMethods.PostMessage(_window, NativeMethods.WmRefresh, IntPtr.Zero, IntPtr.Zero);

        Console.WriteLine("Hook: WH_KEYBOARD_LL, keyDown, head of the hook chain");
        Console.WriteLine("Mapped keys: T R U O C P G M N");
        Console.WriteLine("PowerPoint selection is cached outside the hook. Ctrl+C stops the helper.");

        try
        {
            while (NativeMethods.GetMessage(out NativeMethods.Msg message, IntPtr.Zero, 0, 0) > 0)
                NativeMethods.DispatchMessage(ref message);
        }
        finally
        {
            if (_hook != IntPtr.Zero) NativeMethods.UnhookWindowsHookEx(_hook);
            if (_foregroundHook != IntPtr.Zero) NativeMethods.UnhookWinEvent(_foregroundHook);
            if (_focusHook != IntPtr.Zero) NativeMethods.UnhookWinEvent(_focusHook);
            NativeMethods.KillTimer(_window, new IntPtr(1));
            NativeMethods.DestroyWindow(_window);
            _session.Dispose();
        }

        return 0;
    }

    public static int Describe()
    {
        using var session = new PowerPointSession();
        Snapshot snapshot = session.Peek();
        FocusProbe.Probe probe = FocusProbe.Inspect();
        Console.WriteLine($"surface={snapshot.Surface} {snapshot.Detail}");
        FocusProbe.Write(Console.Out);
        foreach (string token in new[] { "T", "R", "U", "O", "C", "P", "G", "M", "N" })
        {
            GateDecision decision = ShortcutGate.Decide(token, false, false, probe.Focus, snapshot.Surface);
            Console.WriteLine($"{token} {(decision.Dispatch ? "dispatch" : "pass")}");
        }

        return 0;
    }

    IntPtr OnHook(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode != NativeMethods.HcAction) return Next(nCode, wParam, lParam);
        try
        {
            int message = unchecked((int)wParam.ToInt64());
            bool down = message is NativeMethods.WmKeyDown or NativeMethods.WmSysKeyDown;
            bool up = message is NativeMethods.WmKeyUp or NativeMethods.WmSysKeyUp;
            if (!down && !up) return Next(nCode, wParam, lParam);

            var info = Marshal.PtrToStructure<NativeMethods.KbdLlHookStruct>(lParam);
            int vk = (int)info.vkCode;
            if ((info.flags & NativeMethods.LlkhfInjected) != 0) return Next(nCode, wParam, lParam);

            if (up)
            {
                if (vk == _heldVirtualKey)
                {
                    _heldVirtualKey = 0;
                    return Swallow();
                }

                return Next(nCode, wParam, lParam);
            }

            if (!_capture && ShortcutGate.TokenForVirtualKey(vk) == null)
                return Next(nCode, wParam, lParam);

            bool ctrl = IsDown(NativeMethods.VkControl);
            bool shift = IsDown(NativeMethods.VkShift);
            bool alt = IsDown(NativeMethods.VkMenu) || (info.flags & NativeMethods.LlkhfAltDown) != 0;
            bool win = IsDown(NativeMethods.VkLWin) || IsDown(NativeMethods.VkRWin);
            bool caps = (NativeMethods.GetKeyState(NativeMethods.VkCapital) & 1) != 0;

            if (_capture && FocusProbe.Inspect().IsPowerPoint
                && ShortcutGate.TryQueueText(vk, ctrl, alt, win, shift, caps, out char queued))
            {
                if (_pending.Length < QueueLimit) _pending.Append(queued);
                return Swallow();
            }

            if (ctrl || shift || alt || win) return Next(nCode, wParam, lParam);
            string? token = ShortcutGate.TokenForVirtualKey(vk);
            if (token == null) return Next(nCode, wParam, lParam);
            if (vk == _heldVirtualKey) return Swallow();

            FocusProbe.Probe probe = FocusProbe.Inspect();
            if (!probe.IsPowerPoint) return Next(nCode, wParam, lParam);

            GateDecision decision = ShortcutGate.Decide(token, false, false, probe.Focus, _session.EffectiveSurface);
            if (!decision.Dispatch || decision.Token == null)
            {
                NotePass($"{probe.Focus} surface={_session.EffectiveSurface} {_session.Current.Detail}");
                return Next(nCode, wParam, lParam);
            }

            if (!NativeMethods.PostMessage(_window, NativeMethods.WmDispatch, (IntPtr)decision.Token[0], IntPtr.Zero))
                return Next(nCode, wParam, lParam);

            if (decision.Token == "T")
            {
                _session.ArmTextHold();
                _capture = true;
            }

            _heldVirtualKey = vk;
            _lastPass = "";
            return Swallow();
        }
        catch
        {
            return Next(nCode, wParam, lParam);
        }
    }

    IntPtr OnWindow(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam)
    {
        try
        {
            switch (message)
            {
                case NativeMethods.WmDispatch:
                    Dispatch((char)wParam.ToInt32());
                    return IntPtr.Zero;
                case NativeMethods.WmRefresh:
                    _needsRefresh = false;
                    _session.Refresh();
                    WatchPowerPointFocus();
                    return IntPtr.Zero;
                case NativeMethods.WmTimer:
                    if (!_session.IsAttached || _needsRefresh)
                    {
                        _needsRefresh = false;
                        _session.Refresh();
                    }

                    WatchPowerPointFocus();
                    return IntPtr.Zero;
                case NativeMethods.WmClose:
                    NativeMethods.PostQuitMessage(0);
                    return IntPtr.Zero;
            }
        }
        catch
        {
            _capture = false;
        }

        return NativeMethods.DefWindowProc(hwnd, message, wParam, lParam);
    }

    void OnWinEvent(
        IntPtr hook,
        uint eventType,
        IntPtr hwnd,
        int idObject,
        int idChild,
        uint eventThread,
        uint eventTime)
    {
        try
        {
            if (eventType == NativeMethods.EventSystemForeground)
            {
                FocusProbe.Probe probe = FocusProbe.Inspect();
                if (!probe.IsPowerPoint)
                {
                    ReplaceFocusHook(0);
                    return;
                }

                ReplaceFocusHook(probe.ProcessId);
                _session.Invalidate();
                RequestRefresh();
                return;
            }

            if (eventType == NativeMethods.EventObjectFocus)
            {
                if (hwnd == _lastFocus) return;
                _lastFocus = hwnd;
                _session.Invalidate();
                RequestRefresh();
            }
        }
        catch
        {
            _session.Invalidate();
        }
    }

    void Dispatch(char tokenChar)
    {
        string token = tokenChar.ToString();
        if (!ShortcutGate.IsToken(token)) return;
        if (token == "T")
        {
            try
            {
                _session.FinishTextInsert();
                ReplayQueuedText();
            }
            finally
            {
                _capture = false;
            }

            return;
        }

        _session.Dispatch(token);
    }

    void ReplayQueuedText()
    {
        for (int round = 0; round < 6 && _pending.Length > 0; round++)
        {
            string chunk = _pending.Take();
            if (chunk.Length == 0) return;
            if (!SendUnicode(chunk))
            {
                Console.Error.WriteLine("Replay failed");
                return;
            }

            Console.Error.WriteLine($"Replayed {chunk.Length} characters");
        }
    }

    void RequestRefresh()
    {
        _needsRefresh = true;
        if (_window != IntPtr.Zero)
            NativeMethods.PostMessage(_window, NativeMethods.WmRefresh, IntPtr.Zero, IntPtr.Zero);
    }

    void WatchPowerPointFocus()
    {
        FocusProbe.Probe probe = FocusProbe.Inspect();
        ReplaceFocusHook(probe.IsPowerPoint ? probe.ProcessId : 0);
    }

    void ReplaceFocusHook(uint processId)
    {
        if (processId == _focusHookPid) return;
        if (_focusHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWinEvent(_focusHook);
            _focusHook = IntPtr.Zero;
        }

        _focusHookPid = processId;
        if (processId == 0) return;
        _focusHook = NativeMethods.SetWinEventHook(
            NativeMethods.EventObjectFocus,
            NativeMethods.EventObjectFocus,
            IntPtr.Zero,
            _winEventProc,
            processId,
            0,
            NativeMethods.WineventOutOfContext);
    }

    void NotePass(string reason)
    {
        if (reason == _lastPass) return;
        _lastPass = reason;
        Console.Error.WriteLine("pass " + reason);
    }

    IntPtr Next(int nCode, IntPtr wParam, IntPtr lParam) =>
        NativeMethods.CallNextHookEx(_hook, nCode, wParam, lParam);

    static IntPtr Swallow() => new(1);

    static bool IsDown(int virtualKey) => (NativeMethods.GetAsyncKeyState(virtualKey) & 0x8000) != 0;

    static bool SendUnicode(string text)
    {
        var inputs = new NativeMethods.Input[text.Length * 2];
        for (int index = 0; index < text.Length; index++)
        {
            inputs[index * 2] = UnicodeInput(text[index], keyUp: false);
            inputs[index * 2 + 1] = UnicodeInput(text[index], keyUp: true);
        }

        uint sent = NativeMethods.SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<NativeMethods.Input>());
        return sent == (uint)inputs.Length;
    }

    static NativeMethods.Input UnicodeInput(char character, bool keyUp) => new()
    {
        type = NativeMethods.InputKeyboard,
        u = new NativeMethods.InputUnion
        {
            ki = new NativeMethods.KeybdInput
            {
                wVk = 0,
                wScan = character,
                dwFlags = NativeMethods.KeyeventfUnicode | (keyUp ? NativeMethods.KeyeventfKeyUp : 0),
            },
        },
    };

    IntPtr CreateMessageWindow()
    {
        var className = "PPTKKeys.Message." + Environment.ProcessId;
        var windowClass = new NativeMethods.WndClass
        {
            lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_windowProc),
            hInstance = NativeMethods.GetModuleHandle(null),
            lpszClassName = className,
        };
        ushort atom = NativeMethods.RegisterClass(ref windowClass);
        if (atom == 0 && Marshal.GetLastWin32Error() != 1410)
            return IntPtr.Zero;

        return NativeMethods.CreateWindowEx(
            0,
            className,
            "PPTKKeys",
            0,
            0,
            0,
            0,
            0,
            NativeMethods.HwndMessage,
            IntPtr.Zero,
            windowClass.hInstance,
            IntPtr.Zero);
    }
}

internal sealed class StringBuilderPending
{
    readonly System.Text.StringBuilder _builder = new();

    public int Length => _builder.Length;

    public void Append(char character) => _builder.Append(character);

    public string Take()
    {
        string value = _builder.ToString();
        _builder.Clear();
        return value;
    }
}

internal static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        string? command = args.Length > 0 ? args[0] : null;
        if (command is "--help" or "-h" or "/?")
        {
            Console.WriteLine("Usage: PPTKKeys [--describe]");
            Console.WriteLine("Ctrl+C stops the listener.");
            return 0;
        }

        if (command == "--describe") return HelperApp.Describe();
        if (command != null)
        {
            Console.Error.WriteLine("Unknown argument. Usage: PPTKKeys [--describe]");
            return 1;
        }

        using var mutex = new Mutex(false, @"Local\PPTKKeys.Windows");
        bool owned = false;
        try
        {
            try
            {
                owned = mutex.WaitOne(0);
            }
            catch (AbandonedMutexException)
            {
                owned = true;
            }

            if (!owned)
            {
                Console.Error.WriteLine("PPTKKeys is already running.");
                return 1;
            }

            Console.CancelKeyPress += (_, eventArgs) =>
            {
                eventArgs.Cancel = true;
                NativeMethods.PostMessage(MessageWindow.Handle, NativeMethods.WmClose, IntPtr.Zero, IntPtr.Zero);
            };

            return new HelperApp().Run();
        }
        finally
        {
            if (owned) mutex.ReleaseMutex();
        }
    }
}

internal static class MessageWindow
{
    static nint _handle;

    public static IntPtr Handle
    {
        get => (IntPtr)Volatile.Read(ref _handle);
        set => Volatile.Write(ref _handle, (nint)value);
    }
}
