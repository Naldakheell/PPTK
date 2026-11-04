namespace PPTKKeys;

/// <summary>
/// Which virtual keys produce plus, minus, and the batch-2 punctuation on the foreground layout.
/// The scan runs on the helper's message thread. The hook only reads the cached pairs,
/// so it does not call ToUnicode or PowerPoint.
/// US OEM keys are the defaults until the first scan.
/// </summary>
internal static class KeyLayout
{
    readonly record struct Chord(int VirtualKey, bool Shift);

    static readonly Chord[] DefaultPlus = [new(0xBB, true), new(0x6B, false), new(0x6B, true)];
    static readonly Chord[] DefaultMinus = [new(0xBD, false), new(0x6D, false)];
    static readonly Chord[] DefaultSlash = [new(0xBF, false), new(0x6F, false)];
    static readonly Chord[] DefaultPeriod = [new(0xBE, false)];
    static readonly Chord[] DefaultOpenBracket = [new(0xDB, false)];
    static readonly Chord[] DefaultCloseBracket = [new(0xDD, false)];

    static Chord[] _plus = DefaultPlus;
    static Chord[] _minus = DefaultMinus;
    static Chord[] _slash = DefaultSlash;
    static Chord[] _period = DefaultPeriod;
    static Chord[] _openBracket = DefaultOpenBracket;
    static Chord[] _closeBracket = DefaultCloseBracket;
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

    public static bool IsSlash(int virtualKey, bool shift) => Matches(Volatile.Read(ref _slash), virtualKey, shift);

    public static bool IsPeriod(int virtualKey, bool shift) => Matches(Volatile.Read(ref _period), virtualKey, shift);

    public static bool IsOpenBracket(int virtualKey, bool shift) => Matches(Volatile.Read(ref _openBracket), virtualKey, shift);

    public static bool IsCloseBracket(int virtualKey, bool shift) => Matches(Volatile.Read(ref _closeBracket), virtualKey, shift);

    public static bool IsPunctuation(int virtualKey) =>
        Contains(Volatile.Read(ref _slash), virtualKey)
        || Contains(Volatile.Read(ref _period), virtualKey)
        || Contains(Volatile.Read(ref _openBracket), virtualKey)
        || Contains(Volatile.Read(ref _closeBracket), virtualKey);

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
        var slash = new List<Chord>();
        var period = new List<Chord>();
        var openBracket = new List<Chord>();
        var closeBracket = new List<Chord>();
        for (int virtualKey = 0x20; virtualKey <= 0xFE; virtualKey++)
        {
            if (ShortcutSequence.IsModifierVirtualKey(virtualKey)) continue;
            Consider(plus, virtualKey, shift: false, '+', layout);
            Consider(plus, virtualKey, shift: true, '+', layout);
            Consider(minus, virtualKey, shift: false, '-', layout);
            Consider(minus, virtualKey, shift: true, '-', layout);
            Consider(slash, virtualKey, shift: false, '/', layout);
            Consider(slash, virtualKey, shift: true, '/', layout);
            Consider(period, virtualKey, shift: false, '.', layout);
            Consider(period, virtualKey, shift: true, '.', layout);
            Consider(openBracket, virtualKey, shift: false, '[', layout);
            Consider(openBracket, virtualKey, shift: true, '[', layout);
            Consider(closeBracket, virtualKey, shift: false, ']', layout);
            Consider(closeBracket, virtualKey, shift: true, ']', layout);
        }

        bool scannedPlus = plus.Count > 0;
        bool scannedMinus = minus.Count > 0;
        bool scannedSlash = slash.Count > 0;
        Add(plus, 0x6B, false);
        Add(plus, 0x6B, true);
        Add(minus, 0x6D, false);
        Add(slash, 0x6F, false);
        if (!scannedPlus) Add(plus, 0xBB, true);
        if (!scannedMinus) Add(minus, 0xBD, false);
        if (!scannedSlash) Add(slash, 0xBF, false);
        if (period.Count == 0) Add(period, 0xBE, false);
        if (openBracket.Count == 0) Add(openBracket, 0xDB, false);
        if (closeBracket.Count == 0) Add(closeBracket, 0xDD, false);

        Volatile.Write(ref _plus, plus.ToArray());
        Volatile.Write(ref _minus, minus.ToArray());
        Volatile.Write(ref _slash, slash.ToArray());
        Volatile.Write(ref _period, period.ToArray());
        Volatile.Write(ref _openBracket, openBracket.ToArray());
        Volatile.Write(ref _closeBracket, closeBracket.ToArray());
        Console.Error.WriteLine(
            $"layout plus={plus.Count} minus={minus.Count} slash={slash.Count} period={period.Count} open={openBracket.Count} close={closeBracket.Count}");
    }

    static bool Matches(Chord[] table, int virtualKey, bool shift)
    {
        foreach (Chord chord in table)
        {
            if (chord.VirtualKey == virtualKey && chord.Shift == shift) return true;
        }

        return false;
    }

    static bool Contains(Chord[] table, int virtualKey)
    {
        foreach (Chord chord in table)
        {
            if (chord.VirtualKey == virtualKey) return true;
        }

        return false;
    }

    static void Add(List<Chord> found, int virtualKey, bool shift)
    {
        var chord = new Chord(virtualKey, shift);
        if (!found.Contains(chord)) found.Add(chord);
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
