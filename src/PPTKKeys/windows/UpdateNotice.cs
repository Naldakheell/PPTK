using System.Text.Json;
using System.Text.Json.Nodes;

namespace PPTKKeys;

// Phase 2C-1: when to show "A PPTK update is ready. Please close and reopen PowerPoint to complete the update."
// Pure decisions and the "shown" flag in state.json; no window yet (2C-2), not connected to the helper (2C-3).
// The notice never blocks or changes the update itself: the engine applies a pending update when PowerPoint
// closes whether or not the notice was shown.

public static class UpdateNoticeText
{
    public const string Message = "A PPTK update is ready.\nPlease close and reopen PowerPoint\nto complete the update.";
    public static readonly TimeSpan DisplayFor = TimeSpan.FromSeconds(10);
}

/// <summary>What the helper sees at one timer tick. Read on the message thread from data it already has
/// (focus probe, PowerPoint snapshot, time of the last key press); no key contents.</summary>
public sealed record NoticeObservation(bool PowerPointForeground, bool Slideshow, bool DialogOrMenu, TimeSpan SinceLastKey);

public enum NoticeAction
{
    /// <summary>Nothing to announce (no verified pending update, or already shown for this version).</summary>
    None,
    /// <summary>Not now; ask again at the next tick.</summary>
    Wait,
    Show,
    /// <summary>Shown in this session but the flag could not be saved yet: save it, do not show again.</summary>
    MarkShown,
}

public sealed record NoticeDecision(NoticeAction Action, string Reason, string? Version = null);

public static class NotificationPolicy
{
    /// <summary>Typing pause required before the notice may appear.</summary>
    public static readonly TimeSpan IdleBeforeNotice = TimeSpan.FromSeconds(5);

    public static NoticeDecision Decide(PendingInspection pending, NoticeObservation seen)
    {
        if (!pending.HasPending) return new(NoticeAction.None, "no pending update");
        if (!pending.Valid) return new(NoticeAction.None, "the pending update is not verified: " + pending.Problem, pending.Version);
        if (pending.Notified) return new(NoticeAction.None, $"already shown for {pending.Version}", pending.Version);
        if (!seen.PowerPointForeground) return new(NoticeAction.Wait, "PowerPoint is not in the foreground", pending.Version);
        if (seen.Slideshow) return new(NoticeAction.Wait, "a slideshow is running", pending.Version);
        if (seen.DialogOrMenu) return new(NoticeAction.Wait, "a dialog or menu is open", pending.Version);
        if (seen.SinceLastKey < IdleBeforeNotice) return new(NoticeAction.Wait, "the user is typing", pending.Version);
        return new(NoticeAction.Show, "ready", pending.Version);
    }
}

/// <summary>One per helper: makes sure the notice appears at most once per version even when the flag cannot be
/// saved right away (state.lock busy), and records it only after the window actually appeared.</summary>
public sealed class NoticeScheduler
{
    // The version shown in this session, and whether its flag is saved. The inspection the helper decides on may
    // be a few seconds old (it is refreshed by the background worker), so this - not only the saved flag - is
    // what prevents a second display in the same session.
    string? shown;
    bool saved;

    public NoticeDecision Next(PendingInspection pending, NoticeObservation seen)
    {
        if (shown is not null)
        {
            if (pending.HasPending && pending.Version == shown)
            {
                if (!saved && !pending.Notified) return new(NoticeAction.MarkShown, "shown; saving the flag", shown);
                return new(NoticeAction.None, $"already shown for {shown}", shown);
            }
            if (pending.HasPending) { shown = null; saved = false; }   // another version: announce that one
        }
        return NotificationPolicy.Decide(pending, seen);
    }

    /// <summary>The window appeared (in memory only, so it is not shown again in this session).</summary>
    public void Appeared(string version) { shown = version; saved = false; }

    /// <summary>Save the flag for a notice that appeared. Returns what saving did.</summary>
    public MarkResult Save(string version, Func<string, MarkResult> mark) { var result = mark(version); Saved(version, result); return result; }

    /// <summary>Record the result of saving the flag (written elsewhere, e.g. off the message thread).</summary>
    public void Saved(string version, MarkResult result)
    {
        if (shown == version) saved = result is MarkResult.Marked or MarkResult.AlreadyMarked or MarkResult.VersionMismatch or MarkResult.NoPending;
    }

    /// <summary>Call only after the notice window appeared: <see cref="Appeared"/> then <see cref="Save"/>.</summary>
    public MarkResult OnShown(string version, Func<string, MarkResult> mark) { Appeared(version); return Save(version, mark); }
}

/// <summary>Drives the notice from the helper. Inspection results arrive from the background worker
/// (<see cref="OnInspection"/>, any thread); <see cref="Tick"/> runs on the message thread every second and only
/// reads cheap, already available state, shows the window (which takes no focus) and hands the flag write to a
/// background thread. Every failure is logged and swallowed: the notice can never block or break the update
/// engine or the keyboard hook.</summary>
public sealed class UpdateNoticeController(
    Func<NoticeObservation> observe,
    Func<string, bool> show,
    Func<string, MarkResult> mark,
    Action<Action> background,
    Action<string> log)
{
    readonly NoticeScheduler scheduler = new();
    readonly object gate = new();
    PendingInspection? latest;
    string lastReason = "";
    int marking;

    public void OnInspection(PendingInspection inspection) => Volatile.Write(ref latest, inspection);

    public void Tick()
    {
        try
        {
            var pending = Volatile.Read(ref latest);
            // Nothing to announce: skip the observation entirely (it enumerates windows), the answer is None anyway.
            if (pending is null || !pending.HasPending || !pending.Valid || pending.Notified) return;
            NoticeDecision decision;
            lock (gate) decision = scheduler.Next(pending, observe());
            if (decision.Reason != lastReason) { lastReason = decision.Reason; if (decision.Action != NoticeAction.None || pending.HasPending) log($"update notice: {decision.Action} ({decision.Reason})"); }
            switch (decision.Action)
            {
                case NoticeAction.Show:
                    bool appeared;
                    try { appeared = show(decision.Version!); }
                    catch (Exception ex) { log($"update notice: could not be shown ({ex.GetType().Name}: {ex.Message})"); appeared = false; }
                    if (!appeared) { log("update notice: the window did not appear; not recorded, will try again"); return; }
                    lock (gate) scheduler.Appeared(decision.Version!);   // before the next tick can see a stale inspection
                    log($"update notice: shown for {decision.Version}");
                    SaveFlag(decision.Version!);
                    break;
                case NoticeAction.MarkShown:
                    SaveFlag(decision.Version!);
                    break;
            }
        }
        catch (Exception ex) { log($"update notice: tick failed ({ex.GetType().Name}: {ex.Message})"); }
    }

    // One flag write at a time, off the message thread.
    void SaveFlag(string version)
    {
        if (Interlocked.CompareExchange(ref marking, 1, 0) != 0) return;
        try
        {
            background(() =>
            {
                try
                {
                    MarkResult r = mark(version);           // file I/O outside the lock the message thread uses
                    lock (gate) scheduler.Saved(version, r);
                    log($"update notice: flag for {version}: {r}");
                }
                catch (Exception ex) { log($"update notice: flag not saved ({ex.GetType().Name}: {ex.Message})"); }
                finally { Volatile.Write(ref marking, 0); }
            });
        }
        catch { Volatile.Write(ref marking, 0); throw; }
    }
}

public enum MarkResult { Marked, AlreadyMarked, Busy, VersionMismatch, NoPending, Failed }

/// <summary>The "notified" flag of the pending update in state.json, written under state.lock like every other
/// state change. Only the flag changes; install.ps1 keeps it when it stages the same version again.</summary>
public static class PendingNotice
{
    public static MarkResult MarkNotified(PendingUpdatePaths paths, string version)
    {
        FileStream? held;
        try { held = new FileStream(paths.LockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return MarkResult.Busy; }
        using (held)
        {
            try
            {
                if (File.Exists(paths.JournalPath)) return MarkResult.Busy;   // an interrupted operation: leave it to its owner
                if (JsonNode.Parse(File.ReadAllText(paths.StatePath)) is not JsonObject state) return MarkResult.Failed;
                if (state["pending"] is not JsonObject pending) return MarkResult.NoPending;
                if ((string?)pending["version"] != version) return MarkResult.VersionMismatch;
                if (pending["notified"] is JsonValue n && n.GetValueKind() == JsonValueKind.True) return MarkResult.AlreadyMarked;
                pending["notified"] = true;
                pending["notifiedUtc"] = DateTime.UtcNow.ToString("o");
                string tmp = paths.StatePath + ".tmp-" + Guid.NewGuid().ToString("N")[..8];
                File.WriteAllText(tmp, state.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
                try { File.Replace(tmp, paths.StatePath, null); }
                finally { if (File.Exists(tmp)) File.Delete(tmp); }
                return MarkResult.Marked;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return MarkResult.Failed; }
        }
    }
}
