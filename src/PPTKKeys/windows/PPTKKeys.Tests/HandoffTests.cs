using System.Diagnostics;
using System.Text.Json.Nodes;
using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

/// <summary>Handoff protocol with real processes (the HandoffHost test program) and a fake hook. Test mutex and
/// event names only (Local\PPTKKeys.Test.*): the running helper's mutex is never opened or touched.</summary>
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
public class HandoffTests
{
    // prepared, released, mutex, active, recover (ms): short, so failure paths finish quickly.
    const string Timeouts = "6000,6000,1500,6000,10000";

    static string HostDll()
    {
        // ...\PPTKKeys.Tests\bin\<Configuration>\net10.0\ -> ...\PPTKKeys.Tests\HandoffHost\bin\<Configuration>\net10.0\HandoffHost.dll
        var bin = new DirectoryInfo(AppContext.BaseDirectory.TrimEnd('\\'));
        string config = bin.Parent!.Name, tests = bin.Parent!.Parent!.Parent!.FullName;
        var dll = Path.Combine(tests, "HandoffHost", "bin", config, "net10.0", "HandoffHost.dll");
        Assert.True(File.Exists(dll), "HandoffHost not built: " + dll);
        return dll;
    }

    static string Dotnet => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\dotnet\dotnet.exe");

    sealed record LogLine(long Ticks, int Pid, string Role, string What);

    sealed class Run
    {
        public required string Dir { get; init; }
        public required List<LogLine> Lines { get; init; }
        public string? Old => Outcome("old");
        public string? New => Outcome("new");
        public bool StatusActive { get; init; }
        public List<(long Ticks, int Pid)> Killed { get; init; } = [];
        string? Outcome(string role) => Lines.LastOrDefault(l => l.Role == role && l.What.StartsWith("OUTCOME "))?.What["OUTCOME ".Length..];
        public int Installs(string role) => Lines.Count(l => l.Role == role && l.What == "INSTALL");
        public string Dump() => string.Join("\n", Lines.Select(l => $"{l.Ticks} {l.Pid} {l.Role} {l.What}"));

        /// <summary>The hook count over time, in timestamp order: never above one, zero at the end.</summary>
        public void AssertNeverTwoHooks(IEnumerable<(long Ticks, int Pid)>? killed = null)
        {
            var events = Lines.Where(l => l.What is "INSTALL" or "REMOVE" or "DIE").Select(l => (l.Ticks, l.Pid, l.What)).ToList();
            events.AddRange(Lines.Where(l => l.What.StartsWith("KILLED ")).Select(l => (l.Ticks, int.Parse(l.What["KILLED ".Length..]), "KILLED")));
            foreach (var k in (killed ?? []).Concat(Killed).Distinct()) events.Add((k.Ticks, k.Pid, "KILLED"));
            var holding = new HashSet<int>();
            foreach (var e in events.OrderBy(e => e.Ticks))
            {
                if (e.What == "INSTALL") { holding.Add(e.Pid); Assert.True(holding.Count <= 1, "two hooks at once:\n" + Dump()); }
                else holding.Remove(e.Pid);
            }
            Assert.True(holding.Count == 0, "a hook was left installed:\n" + Dump());
        }
    }

    static List<LogLine> ReadLogs(string dir) => Directory.GetFiles(dir, "*.log").SelectMany(f =>
    {
        using var s = new FileStream(f, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var r = new StreamReader(s);
        return r.ReadToEnd().Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(x => x.TrimEnd('\r').Split(' ', 4)).Where(p => p.Length == 4)
            .Select(p => new LogLine(long.Parse(p[0]), int.Parse(p[1]), p[2], p[3])).ToList();
    }).OrderBy(l => l.Ticks).ToList();

    static Process StartOld(string dir, string mutex, string oldBehavior, string newBehavior)
    {
        var psi = new ProcessStartInfo(Dotnet) { UseShellExecute = false, CreateNoWindow = true };
        foreach (var x in new[] { "exec", HostDll(), "--role", "old", "--mutex", mutex, "--log", dir, "--behavior", oldBehavior, "--newbehavior", newBehavior,
            "--timeouts", Timeouts, "--status", Path.Combine(dir, "status.txt"), "--host", HostDll() }) psi.ArgumentList.Add(x);
        return Process.Start(psi)!;
    }

    /// <summary>Runs one handoff. Waits for every host process this run started; kills leftovers and fails if any.</summary>
    static Run Handoff(string oldBehavior = "ok", string newBehavior = "ok", Func<Process, string, (long, int)?>? during = null, List<(long, int)>? killed = null, bool newHangs = false)
    {
        var kills = killed ?? [];
        string dir = Path.Combine(Path.GetTempPath(), "PPTK-handoff-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(dir);
        string mutex = @"Local\PPTKKeys.Test." + Guid.NewGuid().ToString("N")[..12];
        try
        {
            using var a = StartOld(dir, mutex, oldBehavior, newBehavior);
            if (during?.Invoke(a, dir) is { } k) kills.Add(k);
            Assert.True(a.WaitForExit(60_000), "the old host did not finish");
            // The new host (if any) exits after its own linger; wait for it by pid from its log file name.
            foreach (var f in Directory.GetFiles(dir, "*-new.log"))
            {
                int pid = int.Parse(Path.GetFileName(f).Split('-')[0]);
                try
                {
                    using var b = Process.GetProcessById(pid);
                    if (newHangs) { b.Kill(); b.WaitForExit(); kills.Add((Stopwatch.GetTimestamp(), pid)); }   // a stuck new helper, ended by the test
                    else if (!b.WaitForExit(60_000)) { b.Kill(); Assert.Fail("the new host did not finish"); }
                }
                catch (ArgumentException) { }
            }
            return new Run { Dir = dir, Lines = ReadLogs(dir), StatusActive = File.Exists(Path.Combine(dir, "status.txt")) && File.ReadAllText(Path.Combine(dir, "status.txt")) == "active", Killed = kills };
        }
        finally { try { Directory.Delete(dir, true); } catch (IOException) { } }
    }

    [Fact]
    public void Successful_handoff_moves_the_hook_without_overlap()
    {
        var r = Handoff();
        Assert.True(r.Old == "Completed" && r.New == "Active", r.Dump());
        Assert.True(r.StatusActive);
        Assert.Equal(1, r.Installs("old"));
        Assert.Equal(1, r.Installs("new"));
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void New_helper_that_never_gets_ready_leaves_the_old_one_untouched()
    {
        var r = Handoff(newBehavior: "refuse");
        Assert.True(r.Old == "NotPrepared" && r.New == "Refused", r.Dump());
        Assert.Equal(0, r.Installs("new"));
        // The old helper never removed its hook until it shut down.
        Assert.Equal(new[] { "INSTALL", "REMOVE" }, r.Lines.Where(l => l.Role == "old" && l.What is "INSTALL" or "REMOVE").Select(l => l.What));
        Assert.False(r.StatusActive);
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void New_helper_dying_after_release_is_recovered_by_the_old_one()
    {
        var r = Handoff(newBehavior: "die-before-mutex");
        Assert.True(r.Old == "Recovered", r.Dump());
        Assert.Equal(0, r.Installs("new"));
        Assert.Equal(2, r.Installs("old"));
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void New_helper_dying_with_the_hook_installed_is_recovered_from_the_abandoned_mutex()
    {
        var r = Handoff(newBehavior: "die-after-hook");
        Assert.True(r.Old == "Recovered", r.Dump());
        Assert.Contains(r.Lines, l => l.Role == "old" && l.What == "NOTE mutex was abandoned");
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void New_helper_that_cannot_get_the_mutex_exits_and_the_old_one_recovers()
    {
        var r = Handoff(oldBehavior: "squat");
        Assert.True(r.Old == "Recovered" && r.New == "MutexUnavailable", r.Dump());
        Assert.Equal(0, r.Installs("new"));
        r.AssertNeverTwoHooks();
    }

    [Theory]
    [InlineData("hook-fail", "MutexUnavailable")]
    [InlineData("check-fail", "CheckFailed")]
    public void Failed_hook_or_operational_check_steps_down_and_the_old_helper_recovers(string behavior, string outcome)
    {
        var r = Handoff(newBehavior: behavior);
        Assert.True(r.Old == "Recovered" && r.New == outcome, r.Dump());
        Assert.False(r.StatusActive);
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void Abort_after_the_new_helper_took_the_mutex_makes_it_step_down_first()
    {
        // The new helper's operational check takes 8 s; the old one waits 6 s for Active, then aborts and waits for
        // the mutex. The new helper sees the abort before reporting Active and steps down; the old one recovers.
        var r = Handoff(newBehavior: "slow-check");
        Assert.True(r.Old == "Recovered" && r.New == "SteppedDown", r.Dump());
        Assert.Equal(1, r.Installs("new"));
        Assert.False(r.StatusActive);
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void Old_helper_killed_after_releasing_leaves_the_new_one_active()
    {
        var killed = new List<(long, int)>();
        // A slow operational check keeps the new helper busy, so the kill lands between Released and Active.
        var r = Handoff(newBehavior: "slow-check", during: (a, dir) =>
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < 30_000 && !ReadLogs(dir).Any(l => l.Role == "old" && l.What.Contains("mutex released"))) Thread.Sleep(20);
            a.Kill();
            a.WaitForExit();
            return (Stopwatch.GetTimestamp(), a.Id);
        }, killed: killed);
        Assert.True(r.New == "Active", r.Dump());
        Assert.Null(r.Old);   // killed before it could report
        Assert.True(r.StatusActive);
        r.AssertNeverTwoHooks(killed);
    }

    // ---- a new helper that cannot start, or starts and hangs (Phase 2B-3c-1)

    [Fact]
    public void New_helper_dll_that_cannot_load_leaves_the_old_helper_untouched()
    {
        var r = Handoff(oldBehavior: "bad-dll");
        Assert.True(r.Old == "NotPrepared" && r.New is null, r.Dump());
        Assert.Equal(new[] { "INSTALL", "REMOVE" }, r.Lines.Where(l => l.Role == "old" && l.What is "INSTALL" or "REMOVE").Select(l => l.What));   // never released
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void New_helper_hanging_before_it_is_ready_leaves_the_old_helper_untouched()
    {
        var r = Handoff(newBehavior: "hang-before-prepared", newHangs: true);
        Assert.True(r.Old == "NotPrepared", r.Dump());
        Assert.Equal(0, r.Installs("new"));
        Assert.Equal(new[] { "INSTALL", "REMOVE" }, r.Lines.Where(l => l.Role == "old" && l.What is "INSTALL" or "REMOVE").Select(l => l.What));
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void New_helper_hanging_after_release_is_recovered_by_the_old_one()
    {
        var r = Handoff(newBehavior: "hang-after-released", newHangs: true);
        Assert.True(r.Old == "Recovered", r.Dump());
        Assert.Equal(0, r.Installs("new"));
        Assert.Equal(2, r.Installs("old"));
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void New_helper_hanging_with_its_hook_installed_is_not_terminated_while_forced_recovery_is_disabled()
    {
        // The stuck new helper holds the mutex and its hook and ignores Abort. Forced termination is not validated
        // on a real device yet (DeploymentFlags.ForcedRecoveryEnabled = false): the old helper does not end it and
        // does not install a second hook; the hook stays with the new helper, which owns the mutex.
        Assert.False(DeploymentFlags.ForcedRecoveryEnabled);
        var r = Handoff(newBehavior: "hang-in-check", newHangs: true);
        Assert.True(r.Old == "RecoveryFailed", r.Dump());
        Assert.DoesNotContain(r.Lines, l => l.Role == "old" && l.What.StartsWith("KILLED "));
        Assert.Equal(1, r.Installs("new"));
        Assert.Equal(1, r.Installs("old"));
        r.AssertNeverTwoHooks();
    }

    // ---- review findings: window closing mid-handoff, old helper exiting during recovery, late recovery

    [Fact]
    public void Old_helper_window_closing_during_handoff_does_not_crash_and_the_new_one_takes_over()
    {
        // The old helper's window closes just as it would hand the hook over: its finally blocks unhook and release,
        // the handoff call fails (helper thread gone) and is caught; the new helper sees it exit and takes over.
        var r = Handoff(oldBehavior: "close-during-handoff");
        Assert.True(r.Old == "Stopped" && r.New == "Active", r.Dump());
        Assert.Contains(r.Lines, l => l.Role == "old" && l.What.Contains("handoff (old): stopped (HelperThreadGoneException"));
        Assert.True(r.StatusActive);
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void New_helper_window_closing_during_take_over_does_not_crash_and_the_old_one_recovers()
    {
        var r = Handoff(newBehavior: "close-during-takeover");
        Assert.True(r.Old == "Recovered" && r.New == "Stopped", r.Dump());
        Assert.Contains(r.Lines, l => l.Role == "new" && l.What.Contains("handoff (new): stopped (HelperThreadGoneException"));
        Assert.Equal(0, r.Installs("new"));
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void Old_helper_dying_after_abort_does_not_make_the_new_one_step_down()
    {
        // The new helper is slow (8 s check); the old one aborts at 6 s and dies while taking the hook back. The new
        // helper sees the abort but also that the old helper is gone, so it stays: one helper, one hook.
        var r = Handoff(oldBehavior: "die-in-recovery", newBehavior: "slow-check");
        Assert.True(r.New == "Active", r.Dump());
        Assert.Contains(r.Lines, l => l.Role == "new" && l.What.Contains("old helper has exited"));
        Assert.Contains(r.Lines, l => l.Role == "old" && l.What == "DIE");
        Assert.Equal(1, r.Installs("new"));
        Assert.True(r.StatusActive);
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void Old_helper_keeps_trying_when_the_mutex_is_held_past_its_recovery_wait()
    {
        // A third party holds the mutex for 15 s, longer than the 10 s recovery wait: the first recovery fails,
        // the coordinator keeps trying and the old helper gets its hook back.
        var r = Handoff(oldBehavior: "squat-long");
        Assert.True(r.Old == "Recovered", r.Dump());
        Assert.Contains(r.Lines, l => l.Role == "old" && l.What.Contains("recovered late"));
        Assert.Equal(2, r.Installs("old"));
        Assert.Equal(0, r.Installs("new"));
        r.AssertNeverTwoHooks();
    }

    [Fact]
    public void A_new_helper_without_a_handoff_in_progress_does_nothing()
    {
        var names = HandoffNames.ForTest(@"Local\PPTKKeys.Test." + Guid.NewGuid().ToString("N")[..12]);
        int installs = 0;
        var outcome = HelperHandoff.RunNew(names, Environment.ProcessId, HandoffNames.NewNonce(), () => true, new CountingOwner(() => installs++), () => { }, HandoffTimeouts.Default, _ => { });
        Assert.Equal(NewHandoffOutcome.NoHandoff, outcome);
        Assert.Equal(0, installs);
    }

    [Fact]
    public void Pre_created_event_names_are_refused_and_nothing_is_started()
    {
        var names = HandoffNames.ForTest(@"Local\PPTKKeys.Test." + Guid.NewGuid().ToString("N")[..12]);
        string nonce = HandoffNames.NewNonce();
        using var squatted = new EventWaitHandle(false, EventResetMode.ManualReset, names.Event(Environment.ProcessId, nonce, "Prepared"));
        int started = 0;
        var outcome = HelperHandoff.RunOld(names, nonce, () => { started++; return null; }, new CountingOwner(() => { }), HandoffTimeouts.Default, _ => { });
        Assert.Equal(OldHandoffOutcome.StartFailed, outcome);
        Assert.Equal(0, started);
    }

    [Fact]
    public void Test_names_refuse_production_names_and_production_uses_the_helper_mutex()
    {
        Assert.Throws<ArgumentException>(() => HandoffNames.ForTest(HandoffNames.ProductionMutex));
        Assert.Throws<ArgumentException>(() => HandoffNames.ForTest(@"Local\PPTKKeys.Windows.x"));
        Assert.Equal(@"Local\PPTKKeys.Windows", HandoffNames.Production().Mutex);
        Assert.Throws<ArgumentException>(() => HandoffNames.ForTest(@"Local\PPTKKeys.Test.0123abcd").Event(1, "not-hex", "Prepared"));
    }

    [Fact]
    public void Activation_status_is_written_under_the_state_lock()
    {
        using var f = new PendingFixture(keysChanged: true);
        Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);
        using (new FileStream(f.Paths.LockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            Assert.False(HelperActivation.SetStatus(f.Paths, "active"));   // busy: nothing written
        Assert.Null((string?)f.State()["helperActivation"]!["status"]);
        Assert.True(HelperActivation.SetStatus(f.Paths, "active"));
        Assert.Equal("active", (string?)f.State()["helperActivation"]!["status"]);
        Assert.True(HelperActivation.SetStatus(f.Paths, "failed", "the new helper did not become active"));
        Assert.Equal("the new helper did not become active", (string?)f.State()["helperActivation"]!["reason"]);
        Assert.Throws<ArgumentException>(() => HelperActivation.SetStatus(f.Paths, "whatever"));
    }

    [Fact]
    public void No_activation_requested_means_no_status_is_written()
    {
        using var f = new PendingFixture(keysChanged: false);
        Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);
        var before = File.ReadAllText(f.Paths.StatePath);
        Assert.False(HelperActivation.SetStatus(f.Paths, "active"));
        Assert.Equal(before, File.ReadAllText(f.Paths.StatePath));
    }

    sealed class CountingOwner(Action onInstall) : IHookOwner
    {
        public bool AcquireAndInstall(TimeSpan mutexTimeout) { onInstall(); return true; }
        public void RemoveAndRelease() { }
        public bool OperationalCheck() => true;
    }
}
