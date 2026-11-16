using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text.RegularExpressions;
using PPTKKeys;

[assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")]

// Test host for the handoff protocol. Roles:
//   --role old : takes the test mutex and the fake hook, starts "--role new" and hands over (HandoffCoordinator)
//   --role new : takes over (HandoffCoordinator)
// The same coordinator the helper uses. Every hook install/remove is logged with Stopwatch ticks (system-wide)
// to <log>\<pid>-<role>.log, so the tests can prove that two hooks never existed at the same time.
// Fault injection (--behavior / --newbehavior):
//   new side: ok, refuse, die-before-mutex, die-after-hook, hook-fail, check-fail, slow-check,
//             hang-before-prepared, hang-after-released, hang-in-check, close-during-takeover
//   old side: ok, squat, squat-long, bad-dll, close-during-handoff, die-in-recovery

var a = new Dictionary<string, string>();
for (int i = 0; i + 1 < args.Length; i += 2) a[args[i].TrimStart('-')] = args[i + 1];
string role = a["role"], logDir = a["log"];
var names = HandoffNames.ForTest(a["mutex"]);   // refuses production names
var ms = a["timeouts"].Split(',').Select(s => TimeSpan.FromMilliseconds(int.Parse(s))).ToArray();
var timeouts = new HandoffTimeouts(ms[0], ms[1], ms[2], ms[3], ms[4]);
using var log = new HostLog(Path.Combine(logDir, $"{Environment.ProcessId}-{role}.log"), role);
void Note(string m)
{
    log.Write("NOTE " + m);
    var k = Regex.Match(m, @"terminated the unresponsive new helper \(pid (\d+)\)");
    if (k.Success) log.Write("KILLED " + k.Groups[1].Value);   // its hook went with it
}
void Status(string status, string? reason) => File.WriteAllText(a["status"], status);
bool closed = false;

if (role == "old")
{
    using var owner = new FakeHookOwner(names.Mutex, a.GetValueOrDefault("behavior", "ok"), log);
    if (!owner.AcquireAndInstall(TimeSpan.FromSeconds(5))) { log.Write("FAILSTART"); return 10; }
    log.Write("READY");
    var coordinator = new HandoffCoordinator(names, owner, timeouts, Note, Status, () => closed = true, lateRecoveryAttempts: 3);
    Process? StartNew(string nonce)
    {
        // "bad-dll": the new helper's DLL cannot be loaded (missing, or blocked before it runs).
        string target = a.GetValueOrDefault("behavior") == "bad-dll" ? a["host"] + ".missing.dll" : a["host"];
        var psi = new ProcessStartInfo(Environment.ProcessPath!) { UseShellExecute = false, CreateNoWindow = true };
        foreach (var x in new[] { "exec", target, "--role", "new", "--mutex", names.Mutex, "--log", logDir, "--oldpid", Environment.ProcessId.ToString(), "--nonce", nonce,
            "--behavior", a.GetValueOrDefault("newbehavior", "ok"), "--timeouts", a["timeouts"], "--status", a["status"], "--host", a["host"] }) psi.ArgumentList.Add(x);
        return Process.Start(psi);
    }
    var outcome = coordinator.HandOver(StartNew);
    log.Write("OUTCOME " + (outcome?.ToString() ?? "Stopped"));
    if (outcome == OldHandoffOutcome.Completed || outcome is null) return 0;     // B has the hook, or this helper is closing
    Thread.Sleep(int.Parse(a.GetValueOrDefault("linger", "800")));               // A keeps working as the helper
    owner.RemoveAndRelease();
    log.Write("SHUTDOWN");
    return 0;
}
else
{
    using var owner = new FakeHookOwner(names.Mutex, a.GetValueOrDefault("behavior", "ok"), log);
    var coordinator = new HandoffCoordinator(names, owner, timeouts, Note, Status, () => closed = true);
    bool Verify()
    {
        if (a.GetValueOrDefault("behavior") == "hang-before-prepared") { log.Write("NOTE hanging before Prepared"); Thread.Sleep(Timeout.Infinite); }
        return a.GetValueOrDefault("behavior") != "refuse";
    }
    var outcome = coordinator.TakeOver(int.Parse(a["oldpid"]), a["nonce"], Verify, () => { });
    log.Write("OUTCOME " + (outcome?.ToString() ?? "Stopped"));
    if (outcome == NewHandoffOutcome.Active && !closed)
    {
        Thread.Sleep(int.Parse(a.GetValueOrDefault("linger", "800")));            // B works as the helper
        owner.RemoveAndRelease();
        log.Write("SHUTDOWN");
    }
    return 0;
}

sealed class HostLog : IDisposable
{
    readonly StreamWriter writer;
    readonly string role;
    readonly object gate = new();
    public HostLog(string path, string role)
    {
        this.role = role;
        writer = new StreamWriter(new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read)) { AutoFlush = true };
    }
    public void Write(string what) { lock (gate) writer.WriteLine($"{Stopwatch.GetTimestamp()} {Environment.ProcessId} {role} {what}"); }
    public void Dispose() => writer.Dispose();
}

/// <summary>The mutex and a fake hook, both used only from one dedicated thread (mutexes are thread-affine,
/// as on the helper's message thread).</summary>
sealed class FakeHookOwner : IHookOwner, IDisposable
{
    readonly BlockingCollection<Action> work = [];
    readonly Thread thread;
    readonly string mutexName, behavior;
    readonly HostLog log;
    Mutex? mutex;
    bool hook, held, gone;
    int acquireCalls;

    public FakeHookOwner(string mutexName, string behavior, HostLog log)
    {
        this.mutexName = mutexName; this.behavior = behavior; this.log = log;
        thread = new Thread(() => { foreach (var w in work.GetConsumingEnumerable()) w(); }) { IsBackground = true, Name = "hook owner" };
        thread.Start();
    }

    T Run<T>(Func<T> f)
    {
        if (gone) throw new HelperThreadGoneException();   // like SendMessage to a destroyed helper window
        T result = default!; Exception? error = null;
        using var done = new ManualResetEventSlim();
        work.Add(() => { try { result = f(); } catch (Exception e) { error = e; } finally { done.Set(); } });
        done.Wait();
        if (error is not null) throw error;
        return result;
    }

    // What Run's and Main's finally blocks do when the helper window closes: unhook, release, then the thread ends.
    void CloseWindow(string why)
    {
        if (hook) { hook = false; log.Write("REMOVE"); }
        if (held) { held = false; mutex!.ReleaseMutex(); }
        gone = true;
        log.Write("NOTE window closed: " + why);
    }

    public bool AcquireAndInstall(TimeSpan mutexTimeout) => Run(() =>
    {
        acquireCalls++;
        if (behavior == "die-before-mutex") { log.Write("DIE"); Environment.Exit(3); }
        if (behavior == "hang-after-released") { log.Write("NOTE hanging before the mutex"); Thread.Sleep(Timeout.Infinite); }
        if (behavior == "close-during-takeover") { CloseWindow("during take-over"); throw new HelperThreadGoneException(); }
        if (behavior == "die-in-recovery" && acquireCalls > 1) { log.Write("DIE"); Environment.Exit(5); }   // the old helper dies while taking the hook back
        mutex ??= new Mutex(false, mutexName);
        bool got;
        try { got = mutex.WaitOne(mutexTimeout); } catch (AbandonedMutexException) { got = true; log.Write("NOTE mutex was abandoned"); }
        if (!got) { log.Write("NOTE mutex timeout"); return false; }
        held = true;
        if (behavior == "hook-fail") { held = false; mutex.ReleaseMutex(); log.Write("NOTE hook failed"); return false; }
        hook = true;
        log.Write("INSTALL");
        if (behavior == "die-after-hook") { log.Write("DIE"); Environment.Exit(4); }   // the OS removes a dead process's hook; the mutex is abandoned
        return true;
    });

    public void RemoveAndRelease() => Run(() =>
    {
        if (behavior == "close-during-handoff") { CloseWindow("during hand-over"); throw new HelperThreadGoneException(); }
        if (hook) { hook = false; log.Write("REMOVE"); }
        if (held) { held = false; mutex!.ReleaseMutex(); }
        if (behavior is "squat" or "squat-long")
        {
            // Another party grabs the mutex before B can, and holds it for a while (squat-long: past A's recovery wait).
            int holdMs = behavior == "squat" ? 4000 : 15000;
            using var grabbed = new ManualResetEventSlim();
            new Thread(() => { using var m = new Mutex(false, mutexName); m.WaitOne(); grabbed.Set(); Thread.Sleep(holdMs); m.ReleaseMutex(); }) { IsBackground = true }.Start();
            grabbed.Wait();
            log.Write("NOTE squatter holds the mutex");
        }
        return true;
    });

    public bool OperationalCheck() => Run(() =>
    {
        if (behavior == "slow-check") Thread.Sleep(8000);   // longer than the old side's Active wait (6 s in the tests)
        if (behavior == "hang-in-check") { log.Write("NOTE hanging in the operational check with the hook installed"); Thread.Sleep(Timeout.Infinite); }
        return hook && behavior != "check-fail";
    });

    public void Dispose() { work.CompleteAdding(); thread.Join(TimeSpan.FromSeconds(5)); if (!gone) mutex?.Dispose(); }
}
