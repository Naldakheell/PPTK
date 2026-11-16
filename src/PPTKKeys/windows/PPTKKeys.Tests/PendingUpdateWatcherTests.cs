using System.Diagnostics;
using System.Text;
using System.Text.Json.Nodes;
using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

sealed class FakeProbe : IPowerPointProbe
{
    public List<PowerPointProcess> Processes { get; } = [];
    public IReadOnlyList<PowerPointProcess> CurrentSession() { lock (Processes) return Processes.ToList(); }
}

public class PendingUpdateWatcherTests
{
    sealed class Clock { public DateTime Now = new(2026, 11, 14, 12, 0, 0, DateTimeKind.Utc); }

    static (PendingUpdateWatcher Watcher, List<string> Log, FakeProbe Probe, Clock Clock) Watcher(PendingFixture f, Func<PendingAddinUpdate>? engine = null)
    {
        var log = new List<string>(); var probe = new FakeProbe(); var clock = new Clock();
        var w = new PendingUpdateWatcher(f.Paths, engine ?? (() => f.Engine()), probe, m => { lock (log) log.Add(m); }, () => clock.Now);
        return (w, log, probe, clock);
    }

    static async Task RunAsync(PendingUpdateWatcher w, string reason = "test") { Assert.True(w.Trigger(reason)); await w.Current.WaitAsync(TimeSpan.FromSeconds(30)); }
    static Task Done(PendingUpdateWatcher w) => w.Current.WaitAsync(TimeSpan.FromSeconds(30));

    [Fact]
    public async Task Startup_trigger_applies_a_pending_update()
    {
        using var f = new PendingFixture();
        var (w, log, _, _) = Watcher(f);
        await RunAsync(w, "startup");
        Assert.Equal(PendingUpdateOutcome.Applied, w.LastResult!.Outcome);
        Assert.Equal(f.Pending, (string?)f.State()["installed"]);
        Assert.Contains(log, l => l.Contains("pending update (startup): Applied"));
    }

    [Fact]
    public async Task Without_a_pending_folder_the_engine_is_not_run()
    {
        using var f = new PendingFixture();
        Directory.Delete(Path.Combine(f.Paths.PptkDir, "pending"), true);
        var (w, _, _, _) = Watcher(f);
        await RunAsync(w);
        Assert.Equal(0, w.EngineRuns);
    }

    [Fact]
    public async Task Trigger_and_Tick_never_wait_for_the_engine_and_only_one_pass_runs()
    {
        using var f = new PendingFixture();
        using var gate = new ManualResetEventSlim();
        int created = 0;
        var (w, _, _, clock) = Watcher(f, () => { Interlocked.Increment(ref created); gate.Wait(TimeSpan.FromSeconds(30)); return f.Engine(); });

        // Warm-up (first-call JIT is not what the bound is about): a watcher with nothing pending.
        using (var idle = new PendingFixture())
        {
            Directory.Delete(Path.Combine(idle.Paths.PptkDir, "pending"), true);
            var (warm, _, _, _) = Watcher(idle);
            warm.Trigger("warm-up"); warm.Tick(); await Done(warm);
        }

        // The message thread: Trigger/Tick calls while the engine is held.
        var sw = Stopwatch.StartNew();
        Assert.True(w.Trigger("startup"));
        long worst = sw.ElapsedMilliseconds;
        for (int i = 0; i < 200; i++)
        {
            clock.Now = clock.Now.AddSeconds(10);
            var one = Stopwatch.StartNew();
            Assert.False(w.Trigger("again"));
            w.Tick();
            worst = Math.Max(worst, one.ElapsedMilliseconds);
        }
        Assert.True(worst < 50, $"a call on the message thread took {worst} ms");
        gate.Set();
        await Done(w);
        Assert.Equal(1, created);
        Assert.Equal(PendingUpdateOutcome.Applied, w.LastResult!.Outcome);
    }

    [Fact]
    public async Task Tick_starts_at_most_one_pass_per_poll_interval()
    {
        using var f = new PendingFixture();
        Directory.Delete(Path.Combine(f.Paths.PptkDir, "pending"), true);
        var (w, _, probe, clock) = Watcher(f);
        int passes = 0;
        for (int s = 0; s < 20; s++)
        {
            var before = w.Current;
            w.Tick();
            if (!ReferenceEquals(before, w.Current)) passes++;
            await Done(w);
            clock.Now = clock.Now.AddSeconds(1);
        }
        Assert.Equal(4, passes);   // seconds 0, 5, 10, 15
    }

    [Fact]
    public async Task Open_PowerPoint_defers_without_running_the_engine_then_applies_after_it_closes()
    {
        using var f = new PendingFixture();
        var (w, log, probe, _) = Watcher(f);
        probe.Processes.Add(new PowerPointProcess(4242, DateTime.UtcNow, true, "explorer.exe"));
        var before = f.Snapshot();
        await RunAsync(w); await RunAsync(w);
        Assert.Equal(0, w.EngineRuns);
        Assert.Equal(before, f.Snapshot());
        Assert.Single(log, l => l.Contains("PowerPoint is open"));
        probe.Processes.Clear();
        await RunAsync(w);
        Assert.Equal(PendingUpdateOutcome.Applied, w.LastResult!.Outcome);
    }

    [Fact]
    public async Task Hidden_PowerPoint_defers_and_is_logged_once_after_30_seconds_and_never_closed()
    {
        using var f = new PendingFixture();
        var (w, log, probe, clock) = Watcher(f);
        probe.Processes.Add(new PowerPointProcess(9784, new DateTime(2026, 11, 14, 13, 59, 26, DateTimeKind.Utc), false, "explorer.exe"));
        var before = f.Snapshot();
        for (int s = 0; s <= 60; s += 5) { await RunAsync(w); clock.Now = clock.Now.AddSeconds(5); }
        Assert.Equal(0, w.EngineRuns);
        Assert.Equal(before, f.Snapshot());
        var hidden = log.Where(l => l.Contains("without a visible window for")).ToList();
        Assert.Single(hidden);
        Assert.Contains("pid 9784", hidden[0]);
        Assert.Contains("parent explorer.exe", hidden[0]);
        Assert.Contains("left running", hidden[0]);
        Assert.Single(probe.Processes);   // the watcher has no way to act on it
    }

    [Fact]
    public async Task A_short_hidden_phase_while_closing_is_not_logged()
    {
        using var f = new PendingFixture();
        var (w, log, probe, clock) = Watcher(f);
        probe.Processes.Add(new PowerPointProcess(1, DateTime.UtcNow, false, "explorer.exe"));
        await RunAsync(w); clock.Now = clock.Now.AddSeconds(5); await RunAsync(w);
        probe.Processes.Clear(); clock.Now = clock.Now.AddSeconds(5); await RunAsync(w);
        Assert.DoesNotContain(log, l => l.Contains("without a visible window for"));
        Assert.Equal(PendingUpdateOutcome.Applied, w.LastResult!.Outcome);
    }

    [Fact]
    public async Task PowerPoint_starting_after_the_probe_is_caught_by_the_engine()
    {
        using var f = new PendingFixture();
        f.SetProcesses("POWERPNT");   // the engine's own check sees it; the probe did not yet
        var (w, _, _, _) = Watcher(f);
        var before = f.Snapshot();
        await RunAsync(w);
        Assert.Equal(PendingUpdateOutcome.Deferred, w.LastResult!.Outcome);
        Assert.Equal(before, f.Snapshot());
    }

    [Fact]
    public async Task Locked_add_in_defers_then_applies_on_a_later_pass()
    {
        using var f = new PendingFixture();
        var (w, _, _, _) = Watcher(f, () => f.Engine(retries: 1));
        using (new FileStream(f.Paths.AddinFile, FileMode.Open, FileAccess.Read, FileShare.Read)) await RunAsync(w);
        Assert.Equal(PendingUpdateOutcome.Deferred, w.LastResult!.Outcome);
        Assert.Equal(f.Installed, (string?)f.State()["installed"]);
        await RunAsync(w);
        Assert.Equal(PendingUpdateOutcome.Applied, w.LastResult!.Outcome);
    }

    [Fact]
    public async Task Failed_pass_rolls_back_and_a_later_pass_applies()
    {
        using var f = new PendingFixture(keysChanged: true);
        int calls = 0;
        var (w, _, _, _) = Watcher(f, () => ++calls == 1
            ? f.Engine(new PendingUpdateTestHooks { AfterStep = (s, _) => { if (s == "replaced") throw new IOException("simulated"); } })
            : f.Engine());
        var tasks = f.Env.ExportTasks();
        await RunAsync(w);
        Assert.Equal(PendingUpdateOutcome.RolledBack, w.LastResult!.Outcome);
        Assert.Equal(tasks, f.Env.ExportTasks());
        Assert.Equal(PendingFixture.Sha(f.VersionAddin(f.Installed)), PendingFixture.Sha(f.Paths.AddinFile));
        await RunAsync(w);
        Assert.Equal(PendingUpdateOutcome.Applied, w.LastResult!.Outcome);
    }

    [Fact]
    public async Task Interrupted_pass_is_recovered_by_the_next_pass()
    {
        using var f = new PendingFixture();
        int calls = 0;
        var (w, _, _, _) = Watcher(f, () => ++calls == 1 ? f.Engine(new PendingUpdateTestHooks { AbandonAfter = s => s == "replaced" }) : f.Engine());
        await RunAsync(w);
        Assert.Equal(PendingUpdateOutcome.Interrupted, w.LastResult!.Outcome);
        await RunAsync(w);
        Assert.Equal(PendingUpdateOutcome.Applied, w.LastResult!.Outcome);
        Assert.False(File.Exists(f.Paths.JournalPath));
    }

    [Fact]
    public async Task Two_watchers_on_the_same_installation_never_apply_twice()
    {
        using var f = new PendingFixture();
        using var inside = new ManualResetEventSlim(); using var release = new ManualResetEventSlim();
        var (a, _, _, _) = Watcher(f, () => f.Engine(new PendingUpdateTestHooks { AfterStep = (s, _) => { if (s == "journal") { inside.Set(); release.Wait(TimeSpan.FromSeconds(30)); } } }));
        var (b, _, _, _) = Watcher(f);
        Assert.True(a.Trigger("a"));
        Assert.True(inside.Wait(TimeSpan.FromSeconds(30)));
        await RunAsync(b, "b");
        Assert.Equal(PendingUpdateOutcome.Deferred, b.LastResult!.Outcome);   // state.lock
        release.Set(); await Done(a);
        Assert.Equal(PendingUpdateOutcome.Applied, a.LastResult!.Outcome);
        int runs = b.EngineRuns;
        await RunAsync(b, "b again");
        Assert.Equal(runs, b.EngineRuns);   // pending\ is gone: nothing left to run
        Assert.False(Directory.Exists(Path.Combine(f.Paths.PptkDir, "pending")));
        Assert.Equal(f.Pending, (string?)f.State()["installed"]);
    }

    [Fact]
    public async Task An_exception_in_a_pass_is_logged_and_the_next_pass_can_run()
    {
        using var f = new PendingFixture();
        int calls = 0;
        var (w, log, _, _) = Watcher(f, () => ++calls == 1 ? throw new InvalidOperationException("boom") : f.Engine());
        await RunAsync(w);
        Assert.Contains(log, l => l.Contains("pass failed") && l.Contains("boom"));
        await RunAsync(w);
        Assert.Equal(PendingUpdateOutcome.Applied, w.LastResult!.Outcome);
    }
}

/// <summary>Journal contract between the engine (C#) and install.ps1 / uninstall.ps1 (src\Deploy\JOURNAL.md).</summary>
public class JournalContractTests
{
    static readonly UTF8Encoding Utf8 = new(false);

    static Task Done(PendingUpdateWatcher w) => w.Current.WaitAsync(TimeSpan.FromSeconds(30));

    // On failure the message carries the script's output and its own log, so an intermittent failure
    // names the step that failed.
    static void InstallOk(string root, string package)
    {
        var x = PendingAddinUpdateScriptTests.Ps("install.ps1", "-Root", root, "-PackagePath", package);
        Assert.True(x.Code == 0, $"install.ps1 exit {x.Code}: {x.Out} {x.Err}\n{PendingAddinUpdateScriptTests.Logs(root)}");
    }

    [Fact]
    public void Engine_journal_carries_version_and_tasks_format()
    {
        using var f = new PendingFixture();
        Assert.Equal(PendingUpdateOutcome.Interrupted, f.Engine(new PendingUpdateTestHooks { AbandonAfter = s => s == "journal" }).TryApply().Outcome);
        var j = JsonNode.Parse(File.ReadAllText(f.Paths.JournalPath))!;
        Assert.Equal(1, (int)j["journalVersion"]!);
        Assert.Equal("sim-json", (string?)j["tasksFormat"]);
        Assert.Equal(PendingAddinUpdate.Operation, (string?)j["operation"]);
    }

    [Fact]
    public void Engine_refuses_a_journal_in_another_tasks_format()
    {
        using var f = new PendingFixture();
        Assert.Equal(PendingUpdateOutcome.Interrupted, f.Engine(new PendingUpdateTestHooks { AbandonAfter = s => s == "task" }).TryApply().Outcome);
        var j = JsonNode.Parse(File.ReadAllText(f.Paths.JournalPath))!.AsObject();
        j["tasksFormat"] = "task-xml";
        File.WriteAllText(f.Paths.JournalPath, j.ToJsonString());
        var before = f.Snapshot();
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.RestoreFailed, r.Outcome);
        Assert.Contains("task-xml", r.Detail);
        Assert.Equal(before, f.Snapshot());
    }

    [Fact]
    public void Null_tasks_snapshot_removes_only_the_PPTKKeys_task()
    {
        using var f = new PendingFixture();
        var tasks = JsonNode.Parse(File.ReadAllText(Path.Combine(f.Root, "tasks.json")))!.AsArray();
        tasks.Add(new JsonObject { ["name"] = "Backup Sync", ["path"] = "\\", ["actions"] = new JsonArray(new JsonObject { ["execute"] = @"C:\Tools\sync.cmd", ["arguments"] = "" }) });
        File.WriteAllText(Path.Combine(f.Root, "tasks.json"), tasks.ToJsonString(), Utf8);
        f.Env.ImportTasks(null);
        var left = JsonNode.Parse(File.ReadAllText(Path.Combine(f.Root, "tasks.json")))!.AsArray();
        Assert.Single(left);
        Assert.Equal("Backup Sync", (string?)left[0]!["name"]);
    }

    [Fact]
    public void Install_ps1_refuses_a_journal_in_another_tasks_format_and_keeps_it()
    {
        var b = Path.Combine(Path.GetTempPath(), "PPTK-journal-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            var p100 = PendingAddinUpdateScriptTests.Package(b, "1.0.0", "A", "A"); var r = PendingAddinUpdateScriptTests.Root(b);
            InstallOk(r, p100);
            var journal = Path.Combine(r, @"LocalAppData\PPTK\transaction.json");
            File.WriteAllText(journal, new JsonObject { ["journalVersion"] = 1, ["tasksFormat"] = "task-xml", ["operation"] = "apply-pending", ["target"] = "1.0.1", ["tasksBefore"] = "{\"tasks\":[]}" }.ToJsonString(), Utf8);
            var tasks = File.ReadAllText(Path.Combine(r, "tasks.json"));
            var x = PendingAddinUpdateScriptTests.Ps("install.ps1", "-Root", r, "-PackagePath", p100);
            Assert.Equal(1, x.Code);
            Assert.Contains("task-xml", x.Err);
            Assert.True(File.Exists(journal));
            Assert.Equal(tasks, File.ReadAllText(Path.Combine(r, "tasks.json")));
            var u = PendingAddinUpdateScriptTests.Ps("uninstall.ps1", "-Root", r);
            Assert.Equal(1, u.Code);
            Assert.Contains("task-xml", u.Err);
        }
        finally { if (Directory.Exists(b)) Directory.Delete(b, true); }
    }

    [Fact]
    public void Uninstall_ps1_undoes_an_interrupted_engine_run_then_removes_PPTK()
    {
        var b = Path.Combine(Path.GetTempPath(), "PPTK-journal-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            var p100 = PendingAddinUpdateScriptTests.Package(b, "1.0.0", "A", "A"); var p101 = PendingAddinUpdateScriptTests.Package(b, "1.0.1", "B", "B"); var r = PendingAddinUpdateScriptTests.Root(b);
            InstallOk(r, p100);
            File.WriteAllText(Path.Combine(r, "processes.json"), "[\"POWERPNT\"]", Utf8);
            Assert.Equal(1618, PendingAddinUpdateScriptTests.Ps("install.ps1", "-Root", r, "-PackagePath", p101).Code);
            File.WriteAllText(Path.Combine(r, "processes.json"), "[]", Utf8);
            var res = new PendingAddinUpdate(new PendingUpdatePaths(Path.Combine(r, "AppData"), Path.Combine(r, "LocalAppData")), new SimulatedEnvironment(r),
                hooks: new PendingUpdateTestHooks { AbandonAfter = s => s == "replaced" }).TryApply();
            Assert.True(res.Outcome == PendingUpdateOutcome.Interrupted, $"{res.Outcome}: {res.Detail}");

            var u = PendingAddinUpdateScriptTests.Ps("uninstall.ps1", "-Root", r);
            Assert.True(u.Code == 0, u.Out + " " + u.Err);
            Assert.Contains("removed", u.Out);
            var d = PendingAddinUpdateScriptTests.Ps("detect.ps1", "-Root", r, "-ManifestPath", Path.Combine(p100, "version.json"));
            Assert.Contains("NotInstalled", d.Err);
            Assert.False(File.Exists(Path.Combine(r, @"LocalAppData\PPTK\transaction.json")));
        }
        finally { if (Directory.Exists(b)) Directory.Delete(b, true); }
    }

    [Fact]
    public async Task Watcher_applies_an_update_staged_by_install_ps1_and_detect_ps1_confirms_it()
    {
        var b = Path.Combine(Path.GetTempPath(), "PPTK-journal-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            var p100 = PendingAddinUpdateScriptTests.Package(b, "1.0.0", "A", "A"); var p101 = PendingAddinUpdateScriptTests.Package(b, "1.0.1", "B", "A"); var r = PendingAddinUpdateScriptTests.Root(b);
            InstallOk(r, p100);
            File.WriteAllText(Path.Combine(r, "processes.json"), "[\"POWERPNT\"]", Utf8);
            var staged = PendingAddinUpdateScriptTests.Ps("install.ps1", "-Root", r, "-PackagePath", p101);
            Assert.True(staged.Code == 1618, $"install.ps1 exit {staged.Code}: {staged.Out} {staged.Err}\n{PendingAddinUpdateScriptTests.Logs(r)}");

            var paths = new PendingUpdatePaths(Path.Combine(r, "AppData"), Path.Combine(r, "LocalAppData"));
            var env = new SimulatedEnvironment(r);
            var probe = new FakeProbe();
            probe.Processes.Add(new PowerPointProcess(1, DateTime.UtcNow, true, "explorer.exe"));
            var log = new List<string>();
            var w = new PendingUpdateWatcher(paths, () => new PendingAddinUpdate(paths, env), probe, m => { lock (log) log.Add(m); });
            Assert.True(w.Trigger("poll")); await Done(w);
            Assert.Equal(0, w.EngineRuns);   // PowerPoint open: waits
            Assert.Equal(0, PendingAddinUpdateScriptTests.Ps("detect.ps1", "-Root", r, "-ManifestPath", Path.Combine(p100, "version.json")).Code);

            probe.Processes.Clear(); File.WriteAllText(Path.Combine(r, "processes.json"), "[]", Utf8);   // PowerPoint closed
            Assert.True(w.Trigger("poll")); await Done(w);
            Assert.True(w.LastResult!.Outcome == PendingUpdateOutcome.Applied, $"{w.LastResult.Outcome}: {w.LastResult.Detail}\n{string.Join("\n", log)}");
            Assert.False(w.LastResult.HelperRestartRequired);
            var d = PendingAddinUpdateScriptTests.Ps("detect.ps1", "-Root", r, "-ManifestPath", Path.Combine(p101, "version.json"));
            Assert.True(d.Code == 0, d.Err);
            Assert.NotEqual(0, PendingAddinUpdateScriptTests.Ps("detect.ps1", "-Root", r, "-ManifestPath", Path.Combine(p100, "version.json")).Code);
        }
        finally { if (Directory.Exists(b)) Directory.Delete(b, true); }
    }
}
