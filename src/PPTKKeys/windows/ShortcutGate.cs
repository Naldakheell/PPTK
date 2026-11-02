namespace PPTKKeys;

/// <summary>
/// Pure shortcut decision. No Win32 and no COM, so the matrix can be tested on macOS.
/// Sequential and modifier shortcuts are not decided here yet; a later matcher can
/// run before <see cref="Decide"/> without changing the token protocol.
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
        token is "T" or "R" or "U" or "O" or "C" or "P" or "G" or "M" or "N";

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
}
