namespace PPTKKeys;

/// <summary>
/// Pure shortcut decision. No Win32 and no COM, so the matrix can be tested on macOS.
/// Direct inserts go through <see cref="Decide"/>. Two-key canvas sequences go through
/// <see cref="ShortcutSequence"/>. Immediate punctuation and modifier chords go through
/// <see cref="ShortcutChords"/>. None of those paths calls Win32 or PowerPoint.
/// </summary>
public enum Surface
{
    Unknown = 0,
    Canvas = 1,
    Slides = 2,
    Text = 3,
    Blocked = 4,
}

public enum FocusKind
{
    Outside = 0,
    NotFrame = 1,
    MenuOrDialog = 2,
    RibbonOrEdit = 3,
    Uncertain = 4,
    Pane = 5,
}

public readonly record struct GateDecision(bool Swallow, bool Dispatch, string? Token)
{
    public static GateDecision Pass { get; } = new(false, false, null);

    public static GateDecision DispatchToken(string token) => new(true, true, token);
}

public readonly record struct HoldState(bool Held, bool ReleaseAllowed)
{
    public static HoldState Off => new(false, false);

    public static HoldState Armed => new(true, false);

    public static HoldState WaitingForExit => new(true, true);
}

public readonly record struct HoldResult(HoldState Hold, Surface Surface, string Detail);

public static class ShortcutGate
{
    public const string MacroName = "PPTK_DispatchShortcut";

    // Numeric values are the PowerPoint PpSelectionType enum, not declaration order.
    public const int SelectionNone = 0;
    public const int SelectionSlides = 1;
    public const int SelectionShapes = 2;
    public const int SelectionText = 3;

    public const int ViewSlide = 1;
    public const int ViewSlideMaster = 2;
    public const int ViewNotesPage = 3;
    public const int ViewHandoutMaster = 4;
    public const int ViewNotesMaster = 5;
    public const int ViewOutline = 6;
    public const int ViewSlideSorter = 7;
    public const int ViewTitleMaster = 8;
    public const int ViewNormal = 9;
    public const int ViewPrintPreview = 10;
    public const int ViewThumbnails = 11;
    public const int ViewMasterThumbnails = 12;

    public static string PaneName(int pane) => pane switch
    {
        ViewSlide => "slide",
        ViewSlideMaster => "slideMaster",
        ViewNotesPage => "notes",
        ViewHandoutMaster => "handoutMaster",
        ViewNotesMaster => "notesMaster",
        ViewOutline => "outline",
        ViewSlideSorter => "sorter",
        ViewTitleMaster => "titleMaster",
        ViewNormal => "normal",
        ViewPrintPreview => "printPreview",
        ViewThumbnails => "thumbnails",
        ViewMasterThumbnails => "masterThumbnails",
        _ => "other",
    };

    public static string SelectionName(int selection) => selection switch
    {
        SelectionNone => "none",
        SelectionSlides => "slides",
        SelectionShapes => "shapes",
        SelectionText => "text",
        _ => "other",
    };

    public static bool IsToken(string? token) =>
        token is "T" or "R" or "U" or "O" or "C" or "P" or "G" or "M" or "N" or "E" or "Q";

    public static string? TokenForVirtualKey(int virtualKey) => virtualKey switch
    {
        0x54 => "T",
        0x52 => "R",
        0x55 => "U",
        0x4F => "O",
        0x43 => "C",
        0x50 => "P",
        0x47 => "G",
        0x4D => "M",
        0x4E => "N",
        0x45 => "E",
        0x51 => "Q",
        _ => null,
    };

    public static Surface EffectiveSurface(Surface published, bool textHold, bool slideshowBlocked) =>
        slideshowBlocked ? Surface.Blocked : textHold ? Surface.Text : published;

    /// <summary>
    /// Locale-independent pane and selection mapping.
    /// Text wins over every pane. Unknown and blocked both fail open.
    /// Thumbnail and slide sorter allow N only. Outline and notes do not.
    /// </summary>
    public static Surface Classify(bool slideshow, int pane, int selection, bool readOk)
    {
        if (!readOk) return Surface.Unknown;
        if (slideshow) return Surface.Blocked;
        if (selection == SelectionText) return Surface.Text;

        if (pane is ViewNotesPage
            or ViewSlideMaster
            or ViewHandoutMaster
            or ViewNotesMaster
            or ViewOutline
            or ViewTitleMaster
            or ViewPrintPreview
            or ViewMasterThumbnails)
            return Surface.Blocked;

        if (pane is ViewThumbnails or ViewSlideSorter)
            return Surface.Slides;

        if (pane == ViewSlide && selection is SelectionNone or SelectionShapes)
            return Surface.Canvas;

        if (pane == ViewSlide && selection == SelectionSlides)
            return Surface.Slides;

        return Surface.Unknown;
    }

    public static GateDecision Decide(string? token, bool modified, bool injected, FocusKind focus, Surface surface)
    {
        if (modified || injected || !IsToken(token)) return GateDecision.Pass;
        if (focus != FocusKind.Pane) return GateDecision.Pass;
        if (surface == Surface.Canvas) return GateDecision.DispatchToken(token!);
        if (surface == Surface.Slides && token == "N") return GateDecision.DispatchToken(token);
        return GateDecision.Pass;
    }

    public static bool IsTextVirtualKey(int virtualKey) =>
        virtualKey is >= 0x41 and <= 0x5A or >= 0x30 and <= 0x39 or 0x20;

    /// <summary>
    /// Letters typed immediately after T, before the textbox exists, are queued and replayed.
    /// Ctrl, Alt, and the Windows key are not queued.
    /// </summary>
    public static bool TryQueueText(
        int virtualKey,
        bool ctrl,
        bool alt,
        bool win,
        bool shift,
        bool capsLock,
        out char character)
    {
        character = '\0';
        if (ctrl || alt || win) return false;
        if (virtualKey is >= 0x41 and <= 0x5A)
        {
            bool upper = shift ^ capsLock;
            character = (char)(upper ? virtualKey : virtualKey + 32);
            return true;
        }

        if (shift) return false;
        if (virtualKey is >= 0x30 and <= 0x39)
        {
            character = (char)virtualKey;
            return true;
        }

        if (virtualKey == 0x20)
        {
            character = ' ';
            return true;
        }

        return false;
    }

    /// <summary>
    /// After T, bare letters stay disabled until the post-run read sees text,
    /// or a later selection event leaves that text. A plain refresh cannot
    /// publish canvas while the hold is armed, so a stale read cannot re-arm shortcuts.
    /// </summary>
    public static HoldResult Reduce(
        HoldState hold,
        Surface observed,
        string detail,
        bool slideshow,
        bool runFinished,
        bool selectionEvent)
    {
        if (slideshow) return new HoldResult(HoldState.Off, Surface.Blocked, detail);

        if (runFinished)
        {
            if (observed == Surface.Text) return new HoldResult(HoldState.Off, Surface.Text, detail);
            return new HoldResult(HoldState.WaitingForExit, Surface.Text, "hold after T");
        }

        if (!hold.Held) return new HoldResult(HoldState.Off, observed, detail);

        if (observed == Surface.Text) return new HoldResult(HoldState.Off, Surface.Text, detail);

        if (hold.ReleaseAllowed && selectionEvent && observed is Surface.Canvas or Surface.Slides)
            return new HoldResult(HoldState.Off, observed, detail);

        return new HoldResult(hold, Surface.Text, "hold after T");
    }

    public static bool IsIconToken(string? token) => token == "I";

    public static bool IsIconVirtualKey(int virtualKey) => virtualKey == 0x49;

    /// <summary>
    /// I opens Insert Icon only on the slide canvas. Text, ribbon, dialogs,
    /// notes, slideshow, and other panes pass the key through.
    /// </summary>
    public static GateDecision DecideIcon(FocusKind focus, Surface surface)
    {
        if (focus != FocusKind.Pane || surface != Surface.Canvas) return GateDecision.Pass;
        return GateDecision.DispatchToken("I");
    }

    public static bool IsSaveSlidesToken(string? token) => token == "\\";

    /// <summary>
    /// Bare backslash, after the layout has resolved the key to \.
    /// Ctrl, Alt, AltGr, and the Windows key stay with the layout.
    /// </summary>
    public static bool IsBareBackslash(KeyMark mark, bool ctrl, bool alt, bool win) =>
        mark == KeyMark.Backslash && !ctrl && !alt && !win;

    /// <summary>
    /// Save Selected Slides only on the thumbnail pane while slides are selected.
    /// The sorter, the slide canvas, text editing, notes, and a failed read pass the key.
    /// </summary>
    public static GateDecision DecideSaveSlides(FocusKind focus, Surface surface, int pane, int selection)
    {
        if (focus != FocusKind.Pane || surface != Surface.Slides) return GateDecision.Pass;
        if (pane != ViewThumbnails || selection != SelectionSlides) return GateDecision.Pass;
        return GateDecision.DispatchToken("\\");
    }

    public static bool IsRunnable(string? token) =>
        IsToken(token) || IsIconToken(token) || ShortcutSequence.IsSequenceToken(token) || ShortcutChords.IsChordToken(token)
        || IsSaveSlidesToken(token);
}

public enum SequenceSymbol
{
    None = 0,
    Letter = 1,
    Plus = 2,
    Minus = 3,
    Escape = 4,
    Other = 5,
    Period = 6,
}

public enum KeyMark
{
    None = 0,
    Slash = 1,
    Period = 2,
    OpenBracket = 3,
    CloseBracket = 4,
    Backslash = 5,
}

/// <summary>
/// One-key canvas chords. The caller has already resolved punctuation from the
/// keyboard layout, so Ctrl, Alt, and Shift are the keys actually held.
/// AltGr is Ctrl+Alt and is left to the layout. Text editing is the caller's veto.
/// </summary>
public static class ShortcutChords
{
    public static bool IsChordToken(string? token) => token is
        "/" or "CTRL /" or "CTRL SHIFT C" or "CTRL SHIFT V" or "CTRL SHIFT T" or "ALT ." or "CTRL ]" or "CTRL [" or
        "CTRL SHIFT ]" or "CTRL SHIFT [" or
        "CTRL 2" or "CTRL 3" or "ALT 2" or "ALT 3" or
        "CTRL 4" or "CTRL 5" or "CTRL 6" or "CTRL 7" or
        "CTRL K" or "CTRL SPACE";

    public static bool IsModifiedChordKey(int virtualKey) => virtualKey is
        0x32 or 0x33 or 0x34 or 0x35 or 0x36 or 0x37 or
        0x62 or 0x63 or 0x64 or 0x65 or 0x66 or 0x67 or
        0x20 or 0x4B;

    public static string? Token(
        KeyMark mark,
        int virtualKey,
        bool ctrl,
        bool shift,
        bool alt,
        bool win,
        bool unshiftedCloseBracket = false,
        bool unshiftedOpenBracket = false)
    {
        if (mark == KeyMark.Backslash) return null;
        if (win || (ctrl && alt)) return null;

        if (mark == KeyMark.Slash && !alt)
            return ctrl ? "CTRL /" : "/";
        if (mark == KeyMark.Period && alt && !ctrl)
            return "ALT .";
        // A shifted bracket key still types the bracket, so it stays the font command.
        if (mark == KeyMark.CloseBracket && ctrl && !alt)
            return "CTRL ]";
        if (mark == KeyMark.OpenBracket && ctrl && !alt)
            return "CTRL [";
        // US Ctrl+Shift+] holds Shift on the unshifted ] key, which then types }.
        if (ctrl && shift && !alt && unshiftedCloseBracket)
            return "CTRL SHIFT ]";
        if (ctrl && shift && !alt && unshiftedOpenBracket)
            return "CTRL SHIFT [";
        if (mark != KeyMark.None) return null;

        if (virtualKey == 0x43 && ctrl && shift && !alt)
            return "CTRL SHIFT C";
        if (virtualKey == 0x56 && ctrl && shift && !alt)
            return "CTRL SHIFT V";
        if (virtualKey == 0x54 && ctrl && shift && !alt)
            return "CTRL SHIFT T";

        if (shift || win) return null;
        int digit = virtualKey switch
        {
            0x32 or 0x62 => 2,
            0x33 or 0x63 => 3,
            0x34 or 0x64 => 4,
            0x35 or 0x65 => 5,
            0x36 or 0x66 => 6,
            0x37 or 0x67 => 7,
            _ => 0,
        };
        if (digit != 0 && ctrl && !alt) return "CTRL " + digit;
        if ((digit is 2 or 3) && alt && !ctrl) return "ALT " + digit;
        if (virtualKey == 0x4B && ctrl && !alt) return "CTRL K";
        if (virtualKey == 0x20 && ctrl && !alt) return "CTRL SPACE";
        return null;
    }
}

public readonly record struct SequencePress(
    SequenceSymbol Symbol,
    char Letter,
    bool Ctrl,
    bool Alt,
    bool Win,
    bool Shift)
{
    public bool HasCommandModifier => Ctrl || Alt || Win;
}

public readonly record struct SequenceState(char Prefix, uint ArmedAt)
{
    public static SequenceState Idle { get; } = new('\0', 0);

    public bool IsArmed => Prefix != '\0';
}

public readonly record struct SequenceResult(SequenceState State, bool Swallow, string? Token, bool Reprocess);

/// <summary>
/// Two-key sequences. A, S, D, H, V, X, and period arm only when the caller says the slide canvas is legal.
/// The second key must arrive within one second. Escape, an invalid second key, or a lost canvas cancels.
/// Shift+plus and the minus key complete spacing. Shift+minus and shifted letters do not.
/// Period then T merges text. X then T unifies shape type.
/// A shifted period is not the prefix unless that shift is what types the period.
/// </summary>
public static class ShortcutSequence
{
    public const uint TimeoutMs = 1000;

    public static bool IsPrefixLetter(char letter) => letter is 'A' or 'S' or 'D' or 'H' or 'V' or 'X';

    public static bool IsPrefixVirtualKey(int virtualKey) => virtualKey is 0x41 or 0x53 or 0x44 or 0x48 or 0x56 or 0x58;

    public static bool IsModifierVirtualKey(int virtualKey) => virtualKey is
        0x10 or 0x11 or 0x12 or
        0xA0 or 0xA1 or 0xA2 or 0xA3 or 0xA4 or 0xA5 or
        0x5B or 0x5C or
        0x14 or 0x90 or 0x91;

    public static bool IsSequenceToken(string? token) => token is
        "A L" or "A R" or "A T" or "A B" or "A M" or "A C" or "A A" or "A P" or "A X" or "A Z" or
        "S W" or "S H" or "S S" or
        "D H" or "D V" or
        "H +" or "H -" or
        "V +" or "V -" or
        ". T" or "X T";

    public static bool Expired(uint now, uint armedAt) => now - armedAt >= TimeoutMs;

    /// <summary>
    /// Letters use the virtual key, which Windows assigns from the active layout.
    /// Plus and minus use the translated character, with numpad and US OEM fallbacks
    /// when translation returns nothing. Shift+plus is plus. Shift+OEM minus is underscore, not minus.
    /// </summary>
    public static SequencePress PressFromLayout(int virtualKey, bool ctrl, bool alt, bool win, bool shift, char produced)
    {
        if (virtualKey == 0x1B)
            return new SequencePress(SequenceSymbol.Escape, '\0', ctrl, alt, win, shift);
        if (virtualKey == 0x6B)
            return new SequencePress(SequenceSymbol.Plus, '\0', ctrl, alt, win, shift);
        if (virtualKey == 0x6D)
            return new SequencePress(SequenceSymbol.Minus, '\0', ctrl, alt, win, shift);
        if (produced == '+')
            return new SequencePress(SequenceSymbol.Plus, '\0', ctrl, alt, win, shift);
        if (produced == '-')
            return new SequencePress(SequenceSymbol.Minus, '\0', ctrl, alt, win, shift);
        if (produced == '.')
            return new SequencePress(SequenceSymbol.Period, '\0', ctrl, alt, win, shift);
        if (produced == '\0')
        {
            if (virtualKey == 0xBB && shift)
                return new SequencePress(SequenceSymbol.Plus, '\0', ctrl, alt, win, shift);
            if (virtualKey == 0xBD && !shift)
                return new SequencePress(SequenceSymbol.Minus, '\0', ctrl, alt, win, shift);
        }

        if (virtualKey is >= 0x41 and <= 0x5A)
            return new SequencePress(SequenceSymbol.Letter, (char)virtualKey, ctrl, alt, win, shift);
        return new SequencePress(SequenceSymbol.Other, '\0', ctrl, alt, win, shift);
    }

    public static string? Match(char prefix, SequencePress press)
    {
        if (press.HasCommandModifier) return null;
        return (prefix, press.Symbol, press.Letter, press.Shift) switch
        {
            ('A', SequenceSymbol.Letter, 'L', false) => "A L",
            ('A', SequenceSymbol.Letter, 'R', false) => "A R",
            ('A', SequenceSymbol.Letter, 'T', false) => "A T",
            ('A', SequenceSymbol.Letter, 'B', false) => "A B",
            ('A', SequenceSymbol.Letter, 'M', false) => "A M",
            ('A', SequenceSymbol.Letter, 'C', false) => "A C",
            ('A', SequenceSymbol.Letter, 'A', false) => "A A",
            ('A', SequenceSymbol.Letter, 'P', false) => "A P",
            ('A', SequenceSymbol.Letter, 'X', false) => "A X",
            ('A', SequenceSymbol.Letter, 'Z', false) => "A Z",
            ('S', SequenceSymbol.Letter, 'W', false) => "S W",
            ('S', SequenceSymbol.Letter, 'H', false) => "S H",
            ('S', SequenceSymbol.Letter, 'S', false) => "S S",
            ('D', SequenceSymbol.Letter, 'H', false) => "D H",
            ('D', SequenceSymbol.Letter, 'V', false) => "D V",
            ('H', SequenceSymbol.Plus, _, _) => "H +",
            ('H', SequenceSymbol.Minus, _, false) => "H -",
            ('V', SequenceSymbol.Plus, _, _) => "V +",
            ('V', SequenceSymbol.Minus, _, false) => "V -",
            ('X', SequenceSymbol.Letter, 'T', false) => "X T",
            ('.', SequenceSymbol.Letter, 'T', false) => ". T",
            _ => null,
        };
    }

    public static SequenceResult Advance(SequenceState state, SequencePress press, uint now, bool allowArm)
    {
        if (state.IsArmed && Expired(now, state.ArmedAt))
            state = SequenceState.Idle;

        if (state.IsArmed)
        {
            if (!allowArm || press.HasCommandModifier)
                return new SequenceResult(SequenceState.Idle, false, null, true);
            if (press.Symbol == SequenceSymbol.Escape)
                return new SequenceResult(SequenceState.Idle, true, null, false);

            string? token = Match(state.Prefix, press);
            if (token != null)
                return new SequenceResult(SequenceState.Idle, true, token, false);
            return new SequenceResult(SequenceState.Idle, false, null, true);
        }

        if (allowArm
            && !press.HasCommandModifier
            && press.Symbol == SequenceSymbol.Period)
            return new SequenceResult(new SequenceState('.', now), true, null, false);

        if (allowArm
            && press.Symbol == SequenceSymbol.Letter
            && !press.Shift
            && IsPrefixLetter(press.Letter))
            return new SequenceResult(new SequenceState(press.Letter, now), true, null, false);

        return new SequenceResult(SequenceState.Idle, false, null, false);
    }

    /// <summary>
    /// Applies an invalid second key again as a fresh key, so D then D re-arms and A then U can insert.
    /// </summary>
    public static SequenceResult Next(SequenceState state, SequencePress press, uint now, bool allowArm)
    {
        SequenceResult result = Advance(state, press, now, allowArm);
        if (!result.Reprocess) return result;
        return Advance(SequenceState.Idle, press, now, allowArm);
    }
}

/// <summary>
/// Punctuation shortcuts resolved from the keyboard layout. Letter keys are positional
/// shortcuts on every layout, so punctuation never maps onto a letter key. A US punctuation
/// key whose layout types a letter there (Arabic 101: the period key types ز) keeps its
/// US punctuation meaning, the same physical-key rule the letter shortcuts use.
/// </summary>
public static class LayoutRule
{
    public static bool IsLetterKey(int virtualKey) => virtualKey is >= 0x41 and <= 0x5A;

    public static bool KeepsUsPunctuation(char producedUnshifted) => char.IsLetter(producedUnshifted);
}

/// <summary>
/// Right Alt on a layout with AltGr (Arabic 101, most European layouts) reaches the hook as a
/// synthetic Left Ctrl (scan code 0x21D) plus Right Alt. PPTK's Alt chords (Alt+., Alt+2, Alt+3)
/// treat that as plain Alt only when no real Ctrl key is held and the layout types no AltGr
/// character on the key, so AltGr typing and real Ctrl+Alt are left to the layout.
/// </summary>
public static class AltGrRule
{
    public const uint SyntheticCtrlScanCode = 0x21D;

    public static bool IsSyntheticCtrl(int virtualKey, uint scanCode) =>
        virtualKey == 0xA2 && scanCode == SyntheticCtrlScanCode;

    public static bool ActsAsAlt(bool syntheticCtrlHeld, bool rightAltHeld, bool realCtrlHeld, bool keyTypesAltGrCharacter) =>
        syntheticCtrlHeld && rightAltHeld && !realCtrlHeld && !keyTypesAltGrCharacter;
}

/// <summary>
/// PowerPoint (mso20win32client.dll) crashed in QA when Ctrl+2 started immediately after
/// Alt+2 finished (0 ms gap); 30 ms and more never crashed. After a grid command the next
/// queued command waits until GapMs has passed. Nothing else is delayed.
/// </summary>
public static class DispatchPacing
{
    public const uint GapMs = 100;

    public static bool IsGridToken(string? token) => token is "CTRL 2" or "CTRL 3" or "ALT 2" or "ALT 3";

    public static int WaitMs(bool lastWasGrid, uint elapsedSinceLastEnd) =>
        lastWasGrid && elapsedSinceLastEnd < GapMs ? (int)(GapMs - elapsedSinceLastEnd) : 0;
}

public enum RepeatStep
{
    Stop = 0,
    Wait = 1,
    Repeat = 2,
}

/// <summary>
/// Hold-to-repeat for the four spacing sequences only. Windows key auto-repeat drives it:
/// each repeated key-down of the held second key may run the token again. Nothing else repeats.
/// </summary>
public static class ShortcutRepeat
{
    /// <summary>Floor between two runs, so a fast repeat rate cannot overshoot.</summary>
    public const uint MinIntervalMs = 50;

    public static bool IsRepeatable(string? token) => token is "H +" or "H -" or "V +" or "V -";

    /// <summary>
    /// Stop when the token is not repeatable, the canvas is gone, or the held key no longer
    /// completes the same sequence (Shift or a modifier changed). Wait while the previous run
    /// is still queued or running, or inside the minimum interval.
    /// </summary>
    public static RepeatStep Decide(string? heldToken, SequencePress press, bool allowCanvas, bool inFlight, uint now, uint lastRunAt)
    {
        if (!IsRepeatable(heldToken) || !allowCanvas) return RepeatStep.Stop;
        if (ShortcutSequence.Match(heldToken![0], press) != heldToken) return RepeatStep.Stop;
        if (inFlight || now - lastRunAt < MinIntervalMs) return RepeatStep.Wait;
        return RepeatStep.Repeat;
    }
}
