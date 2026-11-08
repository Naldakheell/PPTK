namespace PPTKKeys;

public readonly record struct ReplayOrigin(uint ProcessId, nint FrameWindow);

/// <summary>
/// Foreground observation used to decide whether queued text may be replayed.
/// The keyboard hook only reads this. It does not send input.
/// </summary>
public readonly record struct ReplayObservation(
    bool IsPowerPoint,
    uint ProcessId,
    nint FrameWindow,
    bool FrameClassMatches,
    FocusKind Focus,
    bool MenuOrDialog,
    bool TextEditInFrame);

public enum ReplayDecision
{
    Discard = 0,
    Send = 1,
}

/// <summary>
/// Pure replay policy. Synthetic text is sent only back to the same PowerPoint
/// frame, and only into the slide pane or a text edit inside that frame.
/// </summary>
public static class ReplayGuard
{
    public static bool TryCreateOrigin(
        bool isPowerPoint,
        uint processId,
        nint frameWindow,
        bool frameClassMatches,
        FocusKind focus,
        out ReplayOrigin origin)
    {
        origin = default;
        if (!isPowerPoint || processId == 0 || frameWindow == 0 || !frameClassMatches || focus != FocusKind.Pane)
            return false;
        origin = new ReplayOrigin(processId, frameWindow);
        return true;
    }

    public static bool CanReplay(ReplayOrigin origin, ReplayObservation now, bool powerpointRunning)
    {
        if (!powerpointRunning) return false;
        if (origin.ProcessId == 0 || origin.FrameWindow == 0) return false;
        if (!now.IsPowerPoint || !now.FrameClassMatches) return false;
        if (now.ProcessId != origin.ProcessId || now.FrameWindow != origin.FrameWindow) return false;
        if (now.MenuOrDialog) return false;
        if (now.Focus == FocusKind.Pane) return true;
        return now.Focus == FocusKind.RibbonOrEdit && now.TextEditInFrame;
    }

    public static ReplayDecision Decide(
        ReplayOrigin origin,
        ReplayObservation now,
        bool powerpointRunning,
        int pendingCharacters)
    {
        if (pendingCharacters <= 0) return ReplayDecision.Discard;
        return CanReplay(origin, now, powerpointRunning) ? ReplayDecision.Send : ReplayDecision.Discard;
    }
}
