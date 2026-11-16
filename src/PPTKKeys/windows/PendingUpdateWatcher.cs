namespace PPTKKeys;

// Phase 2B-2: decides when to run PendingAddinUpdate. Not connected to the helper yet, and no live
// environment exists yet (Phase 2B-3): it runs only against injected test environments.
//
// Thread rule: Trigger() and Tick() are called from the helper's message thread (the keyboard hook's
// thread). They never wait: every probe, file check and engine run happens on one background worker.

/// <summary>One POWERPNT.EXE process in the current Windows session. Metadata only: no titles, no paths.</summary>
public sealed record PowerPointProcess(int ProcessId, DateTime StartTimeUtc, bool HasVisibleFrame, string? ParentName);

/// <summary>Read-only view of PowerPoint processes. Tests inject one; the live probe comes with Phase 2B-3.</summary>
public interface IPowerPointProbe
{
    IReadOnlyList<PowerPointProcess> CurrentSession();
}

public sealed class PendingUpdateWatcher
{
    readonly PendingUpdatePaths paths;
    readonly Func<PendingAddinUpdate> createEngine;
    readonly IPowerPointProbe probe;
    readonly Action<string> log;
    readonly Func<DateTime> utcNow;
    readonly TimeSpan pollInterval;
    readonly TimeSpan hiddenThreshold;
    readonly Action<PendingUpdateResult>? onResult;
    readonly Action<PendingInspection>? onInspection;

    int busy;
    long lastPollTicks = long.MinValue;
    Task current = Task.CompletedTask;
    readonly Dictionary<int, DateTime> hiddenSince = [];
    readonly HashSet<int> hiddenReported = [];
    string lastDeferral = "";

    public PendingUpdateWatcher(PendingUpdatePaths paths, Func<PendingAddinUpdate> createEngine, IPowerPointProbe probe, Action<string> log,
        Func<DateTime>? utcNow = null, TimeSpan? pollInterval = null, TimeSpan? hiddenThreshold = null, Action<PendingUpdateResult>? onResult = null,
        Action<PendingInspection>? onInspection = null)
    {
        this.onResult = onResult;
        this.onInspection = onInspection;
        this.paths = paths;
        this.createEngine = createEngine;
        this.probe = probe;
        this.log = log;
        this.utcNow = utcNow ?? (() => DateTime.UtcNow);
        this.pollInterval = pollInterval ?? TimeSpan.FromSeconds(5);
        this.hiddenThreshold = hiddenThreshold ?? TimeSpan.FromSeconds(30);
    }

    /// <summary>The running (or last) background pass. For tests and orderly shutdown.</summary>
    public Task Current => Volatile.Read(ref current);
    public PendingUpdateResult? LastResult { get; private set; }
    public int EngineRuns { get; private set; }

    /// <summary>Starts one background pass unless one is running. Returns at once; false when one was running.</summary>
    public bool Trigger(string reason)
    {
        if (Interlocked.CompareExchange(ref busy, 1, 0) != 0) return false;
        Volatile.Write(ref current, Task.Run(() =>
        {
            try { Pass(reason); }
            catch (Exception ex) { log($"pending update watcher: {reason} pass failed: {ex.GetType().Name}: {ex.Message}"); }
            finally { Volatile.Write(ref busy, 0); }
        }));
        return true;
    }

    /// <summary>For the helper's existing 1-second timer: starts a pass at most once per poll interval.</summary>
    public void Tick()
    {
        long now = utcNow().Ticks;
        long last = Interlocked.Read(ref lastPollTicks);
        if (last != long.MinValue && now - last < pollInterval.Ticks) return;
        if (Trigger("poll")) Interlocked.Exchange(ref lastPollTicks, now);
    }

    void Pass(string reason)
    {
        var processes = probe.CurrentSession();
        NoteHidden(processes);
        if (!Directory.Exists(Path.Combine(paths.PptkDir, "pending")) && !File.Exists(paths.JournalPath)) { Report(NoPending); return; }

        // Any PowerPoint process, visible or not, may hold or reload the add-in: wait. Never close it.
        if (processes.Count > 0)
        {
            string why = processes.Any(p => p.HasVisibleFrame) ? "PowerPoint is open" : "PowerPoint is running without a window";
            if (why != lastDeferral) log($"pending update waits: {why}");
            lastDeferral = why;
            Report(createEngine);   // the update notice is only for this case: PowerPoint open, update waiting
            return;
        }
        lastDeferral = "";

        EngineRuns++;
        var engine = createEngine();
        var result = engine.TryApply();
        LastResult = result;
        foreach (var line in engine.Log) log(line);
        if (result.Outcome != PendingUpdateOutcome.NothingPending) log($"pending update ({reason}): {result.Outcome}: {result.Detail}");
        onResult?.Invoke(result);   // on the worker thread; must not block (the handoff starts its own thread)
        Report(createEngine);       // after the engine, never before it: the notice cannot delay or block an update
    }

    static readonly PendingInspection NoPending = new(false, false, null, false, "no pending update");

    // Read-only inspection for the update notice. Its failures are logged and never reach the engine.
    void Report(PendingInspection inspection)
    {
        if (onInspection is null) return;
        try { onInspection(inspection); } catch (Exception ex) { log($"update notice: inspection not delivered ({ex.GetType().Name}: {ex.Message})"); }
    }

    void Report(Func<PendingAddinUpdate> engine)
    {
        if (onInspection is null) return;
        PendingInspection inspection;
        try { inspection = engine().Inspect(); }
        catch (Exception ex) { log($"update notice: inspection failed ({ex.GetType().Name}: {ex.Message})"); return; }
        Report(inspection);
    }

    // A PowerPoint process without a visible window for longer than the threshold is logged once
    // (pid, start time, parent name). Nothing is done to it.
    void NoteHidden(IReadOnlyList<PowerPointProcess> processes)
    {
        DateTime now = utcNow();
        var present = processes.Select(p => p.ProcessId).ToHashSet();
        foreach (int gone in hiddenSince.Keys.Where(k => !present.Contains(k)).ToList()) { hiddenSince.Remove(gone); hiddenReported.Remove(gone); }
        foreach (var p in processes)
        {
            if (p.HasVisibleFrame) { hiddenSince.Remove(p.ProcessId); hiddenReported.Remove(p.ProcessId); continue; }
            if (!hiddenSince.TryGetValue(p.ProcessId, out var since)) { hiddenSince[p.ProcessId] = now; continue; }
            if (now - since >= hiddenThreshold && hiddenReported.Add(p.ProcessId))
                log($"PowerPoint running without a visible window for {(int)(now - since).TotalSeconds}s: pid {p.ProcessId}, started {p.StartTimeUtc:HH:mm:ss}Z, parent {p.ParentName ?? "unknown"} (left running)");
        }
    }
}
