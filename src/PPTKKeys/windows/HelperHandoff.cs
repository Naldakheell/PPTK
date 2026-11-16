using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Microsoft.Win32.SafeHandles;

namespace PPTKKeys;

// Phase 2B-3b: handing the keyboard hook from the running helper (A) to a newly started one (B).
// Not connected to the helper yet; tested with a host process that has no real hook.
//
// The rule that makes two active hooks impossible: a process installs the hook only while it owns the
// helper mutex, and removes it before releasing the mutex (IHookOwner keeps both together).
//
//   A: start B ................ wait Prepared ─ remove hook, release mutex ─ Released ─ wait Active ─ exit
//   B: check itself ─ Prepared ─ wait Released ─ take mutex, hook, operational check ─ Active ─ run
//   Failures: no Prepared → A keeps everything (Abort). No Active / B died → Abort, A takes the mutex back
//   and reinstalls its hook. B on Abort: removes its hook, releases the mutex, exits.

/// <summary>Mutex and event names. Test names are Local\PPTKKeys.Test.&lt;hex&gt;; production names are refused
/// while production update activation is off.</summary>
public sealed class HandoffNames
{
    public const string ProductionMutex = @"Local\PPTKKeys.Windows";
    static readonly Regex TestMutex = new(@"^Local\\PPTKKeys\.Test\.[0-9a-f]{8,32}$", RegexOptions.CultureInvariant);
    static readonly Regex Nonce = new("^[0-9a-f]{16}$", RegexOptions.CultureInvariant);

    HandoffNames(string mutex) { Mutex = mutex; }

    public string Mutex { get; }

    public static HandoffNames ForTest(string mutex)
    {
        if (!TestMutex.IsMatch(mutex)) throw new ArgumentException($@"test mutex must be Local\PPTKKeys.Test.<hex>, not '{mutex}'");
        return new HandoffNames(mutex);
    }

    public static HandoffNames Production()
    {
        if (!DeploymentFlags.ProductionUpdatesEnabled) throw new InvalidOperationException("production update activation is disabled in this build");
        return new HandoffNames(ProductionMutex);
    }

    public string Event(int oldPid, string nonce, string what)
    {
        if (!Nonce.IsMatch(nonce)) throw new ArgumentException("nonce must be 16 lowercase hex characters");
        return $"{Mutex}.Handoff.{oldPid}.{nonce}.{what}";
    }

    public static string NewNonce() => Convert.ToHexString(RandomNumberGenerator.GetBytes(8)).ToLowerInvariant();
}

/// <summary>The keyboard hook and the mutex that guards it, owned and used by one thread.</summary>
public interface IHookOwner
{
    /// <summary>Take the mutex (an abandoned mutex counts as taken), then install the hook. False when the
    /// mutex was not obtained in time or the hook failed; the mutex is then not held.</summary>
    bool AcquireAndInstall(TimeSpan mutexTimeout);
    /// <summary>Remove the hook, then release the mutex.</summary>
    void RemoveAndRelease();
    /// <summary>True when the hook is installed and the message loop and timer are working.</summary>
    bool OperationalCheck();
}

public sealed record HandoffTimeouts(TimeSpan Prepared, TimeSpan Released, TimeSpan Mutex, TimeSpan Active, TimeSpan Recover)
{
    public static HandoffTimeouts Default { get; } = new(TimeSpan.FromSeconds(20), TimeSpan.FromSeconds(20), TimeSpan.FromSeconds(10), TimeSpan.FromSeconds(20), TimeSpan.FromSeconds(30));
}

public enum OldHandoffOutcome
{
    /// <summary>B is active; A should exit now.</summary>
    Completed,
    /// <summary>B never got ready; A kept its hook the whole time.</summary>
    NotPrepared,
    /// <summary>B did not become active; A took the mutex back and has its hook again.</summary>
    Recovered,
    /// <summary>B could not be started; nothing changed.</summary>
    StartFailed,
    /// <summary>A could not take the mutex back: no hook in this process (logged loudly).</summary>
    RecoveryFailed,
}

public enum NewHandoffOutcome { Active, NoHandoff, Refused, Aborted, MutexUnavailable, HookFailed, CheckFailed, SteppedDown }

[System.Runtime.Versioning.SupportedOSPlatform("windows")]   // named events and mutexes
public static class HelperHandoff
{
    const string Prepared = "Prepared", Released = "Released", ActiveEvent = "Active", Abort = "Abort";

    /// <summary>A's side. Runs on a background thread; the owner does the hook and mutex work on its own thread.</summary>
    public static OldHandoffOutcome RunOld(HandoffNames names, string nonce, Func<Process?> startNew, IHookOwner owner, HandoffTimeouts t, Action<string> log)
    {
        int me = Environment.ProcessId;
        EventWaitHandle? prepared = null, released = null, active = null, abort = null;
        try
        {
            // Created by A before B exists; a name someone else created first is refused (no squatting).
            prepared = CreateNew(names.Event(me, nonce, Prepared)); released = CreateNew(names.Event(me, nonce, Released));
            active = CreateNew(names.Event(me, nonce, ActiveEvent)); abort = CreateNew(names.Event(me, nonce, Abort));
            if (prepared is null || released is null || active is null || abort is null) { log("handoff: event names already in use; not started"); return OldHandoffOutcome.StartFailed; }

            Process? b;
            try { b = startNew(); } catch (Exception ex) { log($"handoff: the new helper could not be started ({ex.GetType().Name}: {ex.Message})"); return OldHandoffOutcome.StartFailed; }
            if (b is null) { log("handoff: the new helper could not be started"); return OldHandoffOutcome.StartFailed; }
            using (b)
            using (var bExit = ExitHandle(b))
            {
                if (WaitHandle.WaitAny([prepared, bExit], t.Prepared) != 0)
                {
                    abort.Set();
                    log("handoff: the new helper did not get ready; keeping the current helper");
                    return OldHandoffOutcome.NotPrepared;
                }
                owner.RemoveAndRelease();
                released.Set();
                log("handoff: hook removed and mutex released for the new helper");

                if (WaitHandle.WaitAny([active, bExit], t.Active) == 0 && !b.HasExited)
                {
                    log($"handoff: the new helper (pid {b.Id}) is active");
                    return OldHandoffOutcome.Completed;
                }
                abort.Set();
                log("handoff: the new helper did not become active; taking the hook back");
                if (owner.AcquireAndInstall(t.Recover)) { log("handoff: recovered, the current helper has its hook again"); return OldHandoffOutcome.Recovered; }
                // B still holds the mutex and ignored Abort for the whole recovery wait: it is unresponsive. A ends
                // the process it started for this handoff (only that one); its hook goes with it and its mutex is
                // abandoned, so A can take both back. Never two hooks: A installs only after taking the mutex.
                if (!DeploymentFlags.ForcedRecoveryEnabled) log("handoff: the new helper did not step down; forced termination is disabled in this build");
                else if (!b.HasExited)
                {
                    log($"handoff: terminated the unresponsive new helper (pid {b.Id})");
                    try { b.Kill(); b.WaitForExit(5000); } catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception) { }
                    if (owner.AcquireAndInstall(t.Recover)) { log("handoff: recovered, the current helper has its hook again"); return OldHandoffOutcome.Recovered; }
                }
                log("handoff: RECOVERY FAILED: the mutex could not be taken back; this helper has no hook");
                return OldHandoffOutcome.RecoveryFailed;
            }
        }
        finally { prepared?.Dispose(); released?.Dispose(); active?.Dispose(); abort?.Dispose(); }
    }

    /// <summary>B's side (started with --handoff &lt;oldPid&gt; &lt;nonce&gt;). On Active, B keeps running as the helper;
    /// it watches for a late Abort until A has exited.</summary>
    public static NewHandoffOutcome RunNew(HandoffNames names, int oldPid, string nonce, Func<bool> verify, IHookOwner owner, Action markActive, HandoffTimeouts t, Action<string> log)
    {
        EventWaitHandle? prepared = Open(names.Event(oldPid, nonce, Prepared)), released = Open(names.Event(oldPid, nonce, Released)),
            active = Open(names.Event(oldPid, nonce, ActiveEvent)), abort = Open(names.Event(oldPid, nonce, Abort));
        try
        {
            if (prepared is null || released is null || active is null || abort is null) { log("handoff: no handoff in progress for this helper"); return NewHandoffOutcome.NoHandoff; }
            using var aExit = OldExitHandle(oldPid);
            if (!verify()) { log("handoff: this helper failed its own checks; not taking over"); return NewHandoffOutcome.Refused; }
            prepared.Set();

            // Abort means "A takes the hook back" - which only works while A is alive. Once A has exited (even right
            // after sending Abort), stepping down would leave no helper at all, so B carries on instead.
            bool OldGone() => aExit.WaitOne(0);
            bool StepDown() => abort.WaitOne(0) && !OldGone();

            // Released, or A gone (its mutex is then released or abandoned): take over. Abort (A alive) or timeout: stop.
            int w = WaitHandle.WaitAny([released, abort, aExit], t.Released);
            if (w == WaitHandle.WaitTimeout || StepDown()) { log("handoff: aborted before taking over"); return NewHandoffOutcome.Aborted; }
            if (abort.WaitOne(0)) log("handoff: aborted, but the old helper has exited; taking over anyway");
            if (!owner.AcquireAndInstall(t.Mutex)) { log("handoff: the mutex or the hook could not be taken; exiting"); return NewHandoffOutcome.MutexUnavailable; }
            if (StepDown()) { owner.RemoveAndRelease(); log("handoff: aborted after taking the mutex; stepped down"); return NewHandoffOutcome.SteppedDown; }
            if (!owner.OperationalCheck())
            {
                if (OldGone()) log("handoff: operational check failed, but the old helper has exited; keeping this hook");
                else { owner.RemoveAndRelease(); log("handoff: operational check failed; stepped down"); return NewHandoffOutcome.CheckFailed; }
            }
            if (StepDown()) { owner.RemoveAndRelease(); log("handoff: aborted during the operational check; stepped down"); return NewHandoffOutcome.SteppedDown; }
            markActive();
            active.Set();
            log("handoff: active");

            // A late Abort (A timed out just before Active) wins while A is alive: step down so A can take the hook back.
            if (WaitHandle.WaitAny([abort, aExit]) == 0 && !OldGone()) { owner.RemoveAndRelease(); log("handoff: aborted after activation; stepped down"); return NewHandoffOutcome.SteppedDown; }
            if (abort.WaitOne(0)) log("handoff: aborted, but the old helper has exited; staying active");
            return NewHandoffOutcome.Active;
        }
        finally { prepared?.Dispose(); released?.Dispose(); active?.Dispose(); abort?.Dispose(); }
    }

    static EventWaitHandle? CreateNew(string name)
    {
        var e = new EventWaitHandle(false, EventResetMode.ManualReset, name, out bool createdNew);
        if (createdNew) return e;
        e.Dispose();
        return null;
    }

    static EventWaitHandle? Open(string name) => EventWaitHandle.TryOpenExisting(name, out var e) ? e : null;

    static WaitHandle ExitHandle(Process p) => new ProcessExitWait(p);

    // A handle signalled when the old helper exits (already signalled if it is gone).
    static WaitHandle OldExitHandle(int pid)
    {
        try
        {
            var p = Process.GetProcessById(pid);
            return new ProcessExitWait(p);
        }
        catch (ArgumentException) { return new ManualResetEvent(true); }
    }

    sealed class ProcessExitWait : WaitHandle
    {
        readonly Process process;
        public ProcessExitWait(Process p) { process = p; SafeWaitHandle = new SafeWaitHandle(p.Handle, ownsHandle: false); }
        protected override void Dispose(bool explicitDisposing) { base.Dispose(explicitDisposing); process.Dispose(); }
    }
}

/// <summary>helperActivation in state.json (written by the update engine): status active / failed, under state.lock.</summary>
public static class HelperActivation
{
    /// <summary>Returns false (nothing written) when the lock is busy or the state is missing or unreadable.</summary>
    public static bool SetStatus(PendingUpdatePaths paths, string status, string? reason = null)
    {
        if (status is not ("active" or "failed")) throw new ArgumentException("status must be active or failed");
        FileStream? held = null;
        try { held = new FileStream(paths.LockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return false; }
        using (held)
        {
            try
            {
                if (JsonNode.Parse(File.ReadAllText(paths.StatePath)) is not JsonObject state) return false;
                if (state["helperActivation"] is not JsonObject activation) return false;   // no activation requested
                activation["status"] = status;
                activation[status == "active" ? "activatedUtc" : "failedUtc"] = DateTime.UtcNow.ToString("o");
                activation["reason"] = reason;
                string tmp = paths.StatePath + ".tmp-" + Guid.NewGuid().ToString("N")[..8];
                File.WriteAllText(tmp, state.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
                File.Replace(tmp, paths.StatePath, null);
                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return false; }
        }
    }
}
