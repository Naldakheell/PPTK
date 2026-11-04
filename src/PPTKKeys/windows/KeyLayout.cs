namespace PPTKKeys;

/// <summary>
/// Which virtual keys produce plus and minus on the foreground keyboard layout.
/// The scan runs on the helper's message thread. The hook only reads the cached pairs,
/// so it does not call ToUnicode or PowerPoint.
/// US Shift+OEM plus and unshifted OEM minus are the defaults until the first scan.
/// </summary>
internal static class KeyLayout
{
    readonly record struct Chord(int VirtualKey, bool Shift);

    static readonly Chord[] DefaultPlus = [new(0xBB, true), new(0x6B, false), new(0x6B, true)];
    static readonly Chord[] DefaultMinus = [new(0xBD, false), new(0x6D, false)];

    static Chord[] _plus = DefaultPlus;
    static Chord[] _minus = DefaultMinus;
    static IntPtr _layout = IntPtr.Zero;

    public static bool IsPlus(int virtualKey, bool shift)
    {
        Chord[] table = Volatile.Read(ref _plus);
        foreach (Chord chord in table)
        {
            if (chord.VirtualKey == virtualKey && chord.Shift == shift) return true;
        }

        return false;
    }

    public static bool IsMinus(int virtualKey, bool shift)
    {
        Chord[] table = Volatile.Read(ref _minus);
        foreach (Chord chord in table)
        {
            if (chord.VirtualKey == virtualKey && chord.Shift == shift) return true;
        }

        return false;
    }

    public static void RefreshForeground()
    {
        uint threadId = 0;
        IntPtr foreground = NativeMethods.GetForegroundWindow();
        if (foreground != IntPtr.Zero)
            threadId = NativeMethods.GetWindowThreadProcessId(foreground, out _);

        IntPtr layout = NativeMethods.GetKeyboardLayout(threadId);
        if (layout == _layout && layout != IntPtr.Zero) return;
        _layout = layout;

        var plus = new List<Chord>();
        var minus = new List<Chord>();
        for (int virtualKey = 0x20; virtualKey <= 0xFE; virtualKey++)
        {
            if (ShortcutSequence.IsModifierVirtualKey(virtualKey)) continue;
            Consider(plus, virtualKey, shift: false, '+', layout);
            Consider(plus, virtualKey, shift: true, '+', layout);
            Consider(minus, virtualKey, shift: false, '-', layout);
            Consider(minus, virtualKey, shift: true, '-', layout);
        }

        bool scannedPlus = plus.Count > 0;
        bool scannedMinus = minus.Count > 0;
        plus.Add(new Chord(0x6B, false));
        plus.Add(new Chord(0x6B, true));
        minus.Add(new Chord(0x6D, false));
        if (!scannedPlus) plus.Add(new Chord(0xBB, true));
        if (!scannedMinus) minus.Add(new Chord(0xBD, false));

        Volatile.Write(ref _plus, plus.ToArray());
        Volatile.Write(ref _minus, minus.ToArray());
        Console.Error.WriteLine($"layout plus={plus.Count} minus={minus.Count}");
    }

    static void Consider(List<Chord> found, int virtualKey, bool shift, char wanted, IntPtr layout)
    {
        if (Produced(virtualKey, shift, layout) != wanted) return;
        var chord = new Chord(virtualKey, shift);
        if (!found.Contains(chord)) found.Add(chord);
    }

    static char Produced(int virtualKey, bool shift, IntPtr layout)
    {
        var state = new byte[256];
        if (shift)
        {
            state[NativeMethods.VkShift] = 0x80;
            state[0xA0] = 0x80;
        }

        state[virtualKey] = 0x80;
        var buffer = new char[8];
        int produced = NativeMethods.ToUnicodeEx((uint)virtualKey, 0, state, buffer, buffer.Length, 0, layout);
        if (produced < 0)
        {
            state[virtualKey] = 0;
            NativeMethods.ToUnicodeEx(NativeMethods.VkSpace, 0, state, buffer, buffer.Length, 0, layout);
            return '\0';
        }

        return produced > 0 ? buffer[0] : '\0';
    }
}
