using System.Runtime.InteropServices;

namespace PPTKKeys;

/// <summary>
/// WH_KEYBOARD_LL helper for the nine PPTK insert tokens and the first canvas sequences.
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
    readonly List<string> _tokens = new();
    SequenceState _sequence = SequenceState.Idle;
    int _heldVirtualKey;
    bool _capture;
    bool _needsRefresh;
    string _lastPass = "";

    const int RefreshTimerId = 1;
    const int SequenceTimerId = 2;

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
        NativeMethods.SetTimer(_window, new IntPtr(RefreshTimerId), 1000, IntPtr.Zero);
        NativeMethods.PostMessage(_window, NativeMethods.WmRefresh, IntPtr.Zero, IntPtr.Zero);

        Console.WriteLine("Hook: WH_KEYBOARD_LL, keyDown, head of the hook chain");
        Console.WriteLine("Mapped keys: T R U O C P G M N E");
        Console.WriteLine("Canvas key: I opens Insert Icon");
        Console.WriteLine("Canvas key: E encloses the selected icon");
        Console.WriteLine("Canvas sequences (1s): A L/R/T/B/M/C/A, S W/H/S, D H/V, H +/-, V +/-, . T");
        Console.WriteLine("Canvas chords: /  Ctrl+/  Ctrl+Shift+C  Ctrl+Shift+V  Ctrl+Shift+T  Alt+.  Ctrl+[  Ctrl+]  Ctrl+Shift+[  Ctrl+Shift+]");
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
            NativeMethods.KillTimer(_window, new IntPtr(RefreshTimerId));
            NativeMethods.KillTimer(_window, new IntPtr(SequenceTimerId));
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
        foreach (string token in new[] { "T", "R", "U", "O", "C", "P", "G", "M", "N", "E" })
        {
            GateDecision decision = ShortcutGate.Decide(token, false, false, probe.Focus, snapshot.Surface);
            Console.WriteLine($"{token} {(decision.Dispatch ? "dispatch" : "pass")}");
        }

        bool sequencesArm = probe.Focus == FocusKind.Pane && snapshot.Surface == Surface.Canvas;
        Console.WriteLine(sequencesArm
            ? "sequences=arm-on-canvas A S D H V ."
            : "sequences=pass");
        Console.WriteLine(sequencesArm
            ? "chords=canvas / Ctrl+/ Ctrl+Shift+C/V Alt+. Ctrl+[/] Ctrl+Shift+[/]"
            : "chords=pass");
        GateDecision saveSlides = ShortcutGate.DecideSaveSlides(
            probe.Focus, snapshot.Surface, snapshot.Pane, snapshot.Selection);
        Console.WriteLine($"save-slides={(saveSlides.Dispatch ? "dispatch" : "pass")}");

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
            if ((info.flags & NativeMethods.LlkhfInjected) != 0)
            {
                if (ShortcutGate.TokenForVirtualKey(vk) != null || ShortcutSequence.IsPrefixVirtualKey(vk))
                    Console.Error.WriteLine($"injected pass vk={vk:X2}");
                return Next(nCode, wParam, lParam);
            }

            if (up)
            {
                if (vk == _heldVirtualKey)
                {
                    _heldVirtualKey = 0;
                    return Swallow();
                }

                return Next(nCode, wParam, lParam);
            }

            if (vk == _heldVirtualKey) return Swallow();
            if (ShortcutSequence.IsModifierVirtualKey(vk)) return Next(nCode, wParam, lParam);

            bool relevant = _capture
                || _sequence.IsArmed
                || ShortcutGate.TokenForVirtualKey(vk) != null
                || ShortcutSequence.IsPrefixVirtualKey(vk)
                || ShortcutGate.IsIconVirtualKey(vk)
                || KeyLayout.IsPunctuation(vk);
            if (!relevant) return Next(nCode, wParam, lParam);

            bool ctrl = IsDown(NativeMethods.VkControl);
            bool shift = IsDown(NativeMethods.VkShift);
            bool alt = IsDown(NativeMethods.VkMenu) || (info.flags & NativeMethods.LlkhfAltDown) != 0;
            bool win = IsDown(NativeMethods.VkLWin) || IsDown(NativeMethods.VkRWin);
            bool caps = (NativeMethods.GetKeyState(NativeMethods.VkCapital) & 1) != 0;

            FocusProbe.Probe probe = FocusProbe.Inspect();
            if (_capture && probe.IsPowerPoint
                && ShortcutGate.TryQueueText(vk, ctrl, alt, win, shift, caps, out char queued))
            {
                if (_pending.Length < QueueLimit) _pending.Append(queued);
                return Swallow();
            }

            if (!probe.IsPowerPoint)
            {
                if (_sequence.IsArmed) CancelSequence("app");
                return Next(nCode, wParam, lParam);
            }

            if (_capture) return Next(nCode, wParam, lParam);

            char produced = ProducedCharacter(vk, shift);
            bool allowCanvas = probe.Focus == FocusKind.Pane
                && _session.EffectiveSurface == Surface.Canvas;
            bool allowArm = allowCanvas && !ctrl && !alt && !win;

            if (_sequence.IsArmed || ShortcutSequence.IsPrefixVirtualKey(vk) || KeyLayout.IsPeriod(vk, shift))
            {
                SequencePress press = ShortcutSequence.PressFromLayout(vk, ctrl, alt, win, shift, produced);
                SequenceResult sequence = ApplySequence(press, info.time, allowArm);
                if (sequence.Token != null)
                {
                    if (!PostToken(sequence.Token)) return Next(nCode, wParam, lParam);
                    _heldVirtualKey = vk;
                    _lastPass = "";
                    return Swallow();
                }

                if (sequence.Swallow)
                {
                    _heldVirtualKey = vk;
                    return Swallow();
                }
            }

            KeyMark mark = KeyMark.None;
            if (KeyLayout.IsSlash(vk, shift)) mark = KeyMark.Slash;
            else if (KeyLayout.IsPeriod(vk, shift)) mark = KeyMark.Period;
            else if (KeyLayout.IsOpenBracket(vk, shift)) mark = KeyMark.OpenBracket;
            else if (KeyLayout.IsCloseBracket(vk, shift)) mark = KeyMark.CloseBracket;
            else if (KeyLayout.IsBackslash(vk, shift)) mark = KeyMark.Backslash;
            if (ShortcutGate.IsBareBackslash(mark, ctrl, alt, win))
            {
                GateDecision saveSlides = ShortcutGate.DecideSaveSlides(
                    probe.Focus,
                    _session.EffectiveSurface,
                    _session.Current.Pane,
                    _session.Current.Selection);
                if (!saveSlides.Dispatch || saveSlides.Token == null) return Next(nCode, wParam, lParam);
                if (!PostToken(saveSlides.Token)) return Next(nCode, wParam, lParam);
                _heldVirtualKey = vk;
                _lastPass = "";
                return Swallow();
            }

            string? chord = ShortcutChords.Token(
                mark,
                vk,
                ctrl,
                shift,
                alt,
                win,
                KeyLayout.IsCloseBracket(vk, false),
                KeyLayout.IsOpenBracket(vk, false));
            if (chord != null)
            {
                bool allowPasteContent = chord == "CTRL SHIFT T"
                    && probe.Focus == FocusKind.Pane
                    && _session.EffectiveSurface == Surface.Text;
                if (!allowCanvas && !allowPasteContent) return Next(nCode, wParam, lParam);
                if (!PostToken(chord)) return Next(nCode, wParam, lParam);
                _heldVirtualKey = vk;
                _lastPass = "";
                return Swallow();
            }

            if (ctrl || shift || alt || win) return Next(nCode, wParam, lParam);
            if (ShortcutGate.IsIconVirtualKey(vk))
            {
                GateDecision icon = ShortcutGate.DecideIcon(probe.Focus, _session.EffectiveSurface);
                if (!icon.Dispatch || icon.Token == null) return Next(nCode, wParam, lParam);
                if (!PostToken(icon.Token)) return Next(nCode, wParam, lParam);
                _heldVirtualKey = vk;
                _lastPass = "";
                return Swallow();
            }

            string? token = ShortcutGate.TokenForVirtualKey(vk);
            if (token == null) return Next(nCode, wParam, lParam);

            GateDecision decision = ShortcutGate.Decide(token, false, false, probe.Focus, _session.EffectiveSurface);
            if (!decision.Dispatch || decision.Token == null)
            {
                NotePass($"{probe.Focus} surface={_session.EffectiveSurface} {_session.Current.Detail}");
                return Next(nCode, wParam, lParam);
            }

            if (!PostToken(decision.Token)) return Next(nCode, wParam, lParam);

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
                    DispatchQueued();
                    return IntPtr.Zero;
                case NativeMethods.WmRefresh:
                    _needsRefresh = false;
                    _session.Refresh();
                    WatchPowerPointFocus();
                    return IntPtr.Zero;
                case NativeMethods.WmTimer:
                    if (wParam == new IntPtr(SequenceTimerId))
                    {
                        if (_sequence.IsArmed
                            && ShortcutSequence.Expired(unchecked((uint)Environment.TickCount), _sequence.ArmedAt))
                            CancelSequence("timeout");
                        return IntPtr.Zero;
                    }

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
            CancelSequence("error");
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
                    CancelSequence("app");
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
                if (_sequence.IsArmed)
                {
                    FocusProbe.Probe focusProbe = FocusProbe.Inspect();
                    if (!focusProbe.IsPowerPoint || focusProbe.Focus != FocusKind.Pane)
                        CancelSequence("focus");
                }

                _session.Invalidate();
                RequestRefresh();
            }
        }
        catch
        {
            _session.Invalidate();
        }
    }

    void DispatchQueued()
    {
        string? token = TakeToken();
        if (token == null || !ShortcutGate.IsRunnable(token)) return;
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

    bool PostToken(string token)
    {
        if (!ShortcutGate.IsRunnable(token) || _tokens.Count >= QueueLimit) return false;
        _tokens.Add(token);
        if (NativeMethods.PostMessage(_window, NativeMethods.WmDispatch, IntPtr.Zero, IntPtr.Zero))
            return true;
        _tokens.RemoveAt(_tokens.Count - 1);
        return false;
    }

    string? TakeToken()
    {
        if (_tokens.Count == 0) return null;
        string token = _tokens[0];
        _tokens.RemoveAt(0);
        return token;
    }

    SequenceResult ApplySequence(SequencePress press, uint now, bool allowArm)
    {
        bool wasArmed = _sequence.IsArmed;
        uint armedAt = _sequence.ArmedAt;
        bool timedOut = wasArmed && ShortcutSequence.Expired(now, armedAt);
        if (timedOut) Console.Error.WriteLine("Sequence timeout");

        SequenceResult result = ShortcutSequence.Next(_sequence, press, now, allowArm);
        if (result.State.IsArmed && result.State.ArmedAt != armedAt)
        {
            NativeMethods.SetTimer(_window, new IntPtr(SequenceTimerId), ShortcutSequence.TimeoutMs, IntPtr.Zero);
            Console.Error.WriteLine("Sequence armed " + result.State.Prefix);
        }
        else if (wasArmed && !result.State.IsArmed)
        {
            NativeMethods.KillTimer(_window, new IntPtr(SequenceTimerId));
            if (!timedOut && result.Token == null)
                Console.Error.WriteLine(result.Swallow ? "Sequence cancelled escape" : "Sequence cancelled key");
        }

        _sequence = result.State;
        return result;
    }

    void CancelSequence(string reason)
    {
        if (!_sequence.IsArmed) return;
        _sequence = SequenceState.Idle;
        if (_window != IntPtr.Zero)
            NativeMethods.KillTimer(_window, new IntPtr(SequenceTimerId));
        Console.Error.WriteLine("Sequence cancelled " + reason);
    }

    static char ProducedCharacter(int virtualKey, bool shift)
    {
        if (virtualKey is >= 0x41 and <= 0x5A or 0x1B) return '\0';
        if (KeyLayout.IsPlus(virtualKey, shift)) return '+';
        if (KeyLayout.IsMinus(virtualKey, shift)) return '-';
        if (KeyLayout.IsPeriod(virtualKey, shift)) return '.';
        return '\0';
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
        KeyLayout.RefreshForeground();
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
