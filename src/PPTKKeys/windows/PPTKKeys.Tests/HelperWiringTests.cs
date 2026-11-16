using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.Versioning;
using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

/// <summary>A dedicated "message thread" for HelperHookOwner: work runs there and the caller waits (as SendMessage does).</summary>
sealed class OwnerThread : IDisposable
{
    readonly BlockingCollection<Action> work = [];
    readonly Thread thread;
    public int Id { get; }

    public OwnerThread()
    {
        using var started = new ManualResetEventSlim();
        int id = 0;
        thread = new Thread(() => { id = Environment.CurrentManagedThreadId; started.Set(); foreach (var w in work.GetConsumingEnumerable()) w(); }) { IsBackground = true };
        thread.Start();
        started.Wait();
        Id = id;
    }

    public void Run(Action action)
    {
        if (Environment.CurrentManagedThreadId == Id) { action(); return; }
        Exception? error = null;
        using var done = new ManualResetEventSlim();
        work.Add(() => { try { action(); } catch (Exception e) { error = e; } finally { done.Set(); } });
        done.Wait();
        if (error is not null) throw error;
    }

    public void Dispose() { work.CompleteAdding(); thread.Join(5000); }
}

[SupportedOSPlatform("windows")]
public class HelperWiringTests
{
    static string TestMutex() => @"Local\PPTKKeys.Test." + Guid.NewGuid().ToString("N")[..12];

    sealed class FakeHook
    {
        public readonly List<string> Calls = [];
        public int InstallThread, UninstallThread;
        public bool Fail;
        public IntPtr Install(Func<bool> mutexHeldElsewhere)
        {
            InstallThread = Environment.CurrentManagedThreadId;
            Calls.Add(mutexHeldElsewhere() ? "install(mutex free!)" : "install");
            return Fail ? IntPtr.Zero : new IntPtr(42);
        }
        public void Uninstall(IntPtr h) { UninstallThread = Environment.CurrentManagedThreadId; Calls.Add("uninstall " + h); }
    }

    // True when another thread can take the mutex right now (i.e. this owner does NOT hold it).
    static bool FreeForOthers(string name)
    {
        bool free = false;
        var t = new Thread(() =>
        {
            using var m = new Mutex(false, name);
            try { free = m.WaitOne(0); } catch (AbandonedMutexException) { free = true; }
            if (free) m.ReleaseMutex();
        });
        t.Start(); t.Join();
        return free;
    }

    [Fact]
    public void Owner_takes_the_mutex_before_installing_and_removes_the_hook_before_releasing()
    {
        string name = TestMutex();
        using var thread = new OwnerThread();
        Mutex? mutex = null; thread.Run(() => mutex = new Mutex(false, name));
        var hook = new FakeHook();
        long ticks = 0;
        var owner = new HelperHookOwner(mutex!, false, () => hook.Install(() => FreeForOthers(name)), hook.Uninstall, () => Interlocked.Read(ref ticks), thread.Run);

        Assert.True(owner.AcquireAndInstall(TimeSpan.FromSeconds(1)));   // called from the test thread
        Assert.Equal(new[] { "install" }, hook.Calls);                     // the mutex was held during install
        Assert.Equal(thread.Id, hook.InstallThread);                        // on the owner thread
        Assert.False(FreeForOthers(name));
        Assert.True(owner.Held);

        owner.RemoveAndRelease();
        Assert.Equal("uninstall 42", hook.Calls[^1]);
        Assert.Equal(thread.Id, hook.UninstallThread);
        Assert.True(FreeForOthers(name));
        Assert.False(owner.Held);
        thread.Run(() => mutex!.Dispose());
    }

    [Fact]
    public void Failed_install_releases_the_mutex()
    {
        string name = TestMutex();
        using var thread = new OwnerThread();
        Mutex? mutex = null; thread.Run(() => mutex = new Mutex(false, name));
        var hook = new FakeHook { Fail = true };
        var owner = new HelperHookOwner(mutex!, false, () => hook.Install(() => false), hook.Uninstall, () => 0, thread.Run);
        Assert.False(owner.AcquireAndInstall(TimeSpan.FromSeconds(1)));
        Assert.False(owner.Held);
        Assert.True(FreeForOthers(name));
        thread.Run(() => mutex!.Dispose());
    }

    [Fact]
    public void Mutex_held_elsewhere_times_out_without_installing()
    {
        string name = TestMutex();
        using var other = new OwnerThread();
        Mutex? held = null; other.Run(() => { held = new Mutex(false, name); held.WaitOne(); });
        using var thread = new OwnerThread();
        Mutex? mutex = null; thread.Run(() => mutex = new Mutex(false, name));
        var hook = new FakeHook();
        var owner = new HelperHookOwner(mutex!, false, () => hook.Install(() => false), hook.Uninstall, () => 0, thread.Run);
        Assert.False(owner.AcquireAndInstall(TimeSpan.FromMilliseconds(200)));
        Assert.Empty(hook.Calls);
        other.Run(() => { held!.ReleaseMutex(); held.Dispose(); });
        thread.Run(() => mutex!.Dispose());
    }

    [Fact]
    public void Abandoned_mutex_counts_as_taken()
    {
        string name = TestMutex();
        var dead = new Thread(() => { var m = new Mutex(false, name); m.WaitOne(); });   // exits without releasing
        dead.Start(); dead.Join();
        using var thread = new OwnerThread();
        Mutex? mutex = null; thread.Run(() => mutex = new Mutex(false, name));
        var hook = new FakeHook();
        var owner = new HelperHookOwner(mutex!, false, () => hook.Install(() => false), hook.Uninstall, () => 0, thread.Run);
        Assert.True(owner.AcquireAndInstall(TimeSpan.FromSeconds(1)));
        owner.RemoveAndRelease();
        thread.Run(() => mutex!.Dispose());
    }

    [Fact]
    public void Adopted_hook_and_mutex_from_a_normal_start_are_handed_over_the_same_way()
    {
        string name = TestMutex();
        using var thread = new OwnerThread();
        Mutex? mutex = null; thread.Run(() => { mutex = new Mutex(false, name); mutex.WaitOne(0); });   // as Main does
        var hook = new FakeHook();
        var owner = new HelperHookOwner(mutex!, true, () => hook.Install(() => false), hook.Uninstall, () => 0, thread.Run);
        owner.Adopt(new IntPtr(7));
        Assert.True(owner.Held);
        owner.RemoveAndRelease();
        Assert.Equal(new[] { "uninstall 7" }, hook.Calls);
        Assert.True(FreeForOthers(name));
        thread.Run(() => mutex!.Dispose());
    }

    [Fact]
    public void Operational_check_needs_the_hook_and_a_ticking_timer()
    {
        string name = TestMutex();
        using var thread = new OwnerThread();
        Mutex? mutex = null; thread.Run(() => mutex = new Mutex(false, name));
        var hook = new FakeHook();
        long ticks = 0;
        var owner = new HelperHookOwner(mutex!, false, () => hook.Install(() => false), hook.Uninstall, () => Interlocked.Read(ref ticks), thread.Run, TimeSpan.FromMilliseconds(400));
        Assert.False(owner.OperationalCheck());   // no hook
        Assert.True(owner.AcquireAndInstall(TimeSpan.FromSeconds(1)));
        Assert.False(owner.OperationalCheck());   // timer not ticking
        using var timer = new Timer(_ => Interlocked.Increment(ref ticks), null, 0, 50);
        Assert.True(owner.OperationalCheck());
        owner.RemoveAndRelease();
        thread.Run(() => mutex!.Dispose());
    }

    // ---- review fix 1: a normal start whose hook install fails still releases the mutex

    [Fact]
    public void Normal_start_with_a_failed_hook_install_still_releases_the_mutex()
    {
        string name = TestMutex();
        using var thread = new OwnerThread();
        Mutex? mutex = null; bool owned = false;
        thread.Run(() => { mutex = new Mutex(false, name); owned = mutex.WaitOne(0); });   // Main: WaitOne(0)
        Assert.True(owned);
        var hook = new FakeHook();
        // HelperApp.Run creates the owner with the mutex already held; the hook install then fails, so Adopt never runs.
        var owner = new HelperHookOwner(mutex!, owned, () => hook.Install(() => false), hook.Uninstall, () => 0, thread.Run);
        Assert.True(owner.Held);                     // = HelperApp.MutexHeld
        Assert.Equal(IntPtr.Zero, owner.Hook);
        thread.Run(() => { if (owner.Held) mutex!.ReleaseMutex(); });   // Main's finally, on the helper thread
        Assert.True(FreeForOthers(name));            // released, not abandoned
        thread.Run(() => mutex!.Dispose());
    }

    // ---- review fix 2: work for a helper thread that has gone

    [Fact]
    public void Calls_after_the_helper_thread_has_gone_raise_HelperThreadGone()
    {
        string name = TestMutex();
        using var thread = new OwnerThread();
        Mutex? mutex = null; thread.Run(() => mutex = new Mutex(false, name));
        var hook = new FakeHook();
        bool gone = false;
        var owner = new HelperHookOwner(mutex!, false, () => hook.Install(() => false), hook.Uninstall, () => 0,
            work => { if (gone) throw new HelperThreadGoneException(); thread.Run(work); });
        Assert.True(owner.AcquireAndInstall(TimeSpan.FromSeconds(1)));
        gone = true;
        Assert.Throws<HelperThreadGoneException>(() => owner.RemoveAndRelease());
        Assert.Throws<HelperThreadGoneException>(() => owner.AcquireAndInstall(TimeSpan.Zero));
        gone = false; owner.RemoveAndRelease();
        thread.Run(() => mutex!.Dispose());
    }

    [Fact]
    public void Coordinator_never_throws_when_the_helper_thread_has_gone()
    {
        var names = HandoffNames.ForTest(TestMutex());
        var log = new List<string>(); var status = new List<string>(); int closes = 0;
        var c = new HandoffCoordinator(names, new ThrowingOwner(), HandoffTimeouts.Default, log.Add, (s, r) => status.Add(s), () => closes++);
        // Old side: the start function fails as a closing helper would.
        Assert.Equal(OldHandoffOutcome.StartFailed, c.HandOver(_ => throw new HelperThreadGoneException()));
        // New side with no handoff in progress: NoHandoff, closes; nothing thrown.
        Assert.Equal(NewHandoffOutcome.NoHandoff, c.TakeOver(Environment.ProcessId, HandoffNames.NewNonce(), () => true, () => { }));
        Assert.Equal(1, closes);
        Assert.Contains("failed", status);
    }

    sealed class ThrowingOwner : IHookOwner
    {
        public bool AcquireAndInstall(TimeSpan mutexTimeout) => throw new HelperThreadGoneException();
        public void RemoveAndRelease() => throw new HelperThreadGoneException();
        public bool OperationalCheck() => throw new HelperThreadGoneException();
    }

    // ---- review fix 6: operational check waits long enough for a slow timer, without blocking the owner thread

    [Fact]
    public async Task Operational_check_tolerates_a_slow_timer_and_does_not_block_the_owner_thread()
    {
        string name = TestMutex();
        using var thread = new OwnerThread();
        Mutex? mutex = null; thread.Run(() => mutex = new Mutex(false, name));
        var hook = new FakeHook();
        long ticks = 0;
        var owner = new HelperHookOwner(mutex!, false, () => hook.Install(() => false), hook.Uninstall, () => Interlocked.Read(ref ticks), thread.Run);   // default wait (5 s)
        Assert.True(owner.AcquireAndInstall(TimeSpan.FromSeconds(1)));
        using var slowTimer = new Timer(_ => Interlocked.Increment(ref ticks), null, 2500, 2500);   // a refresh that takes 2.5 s
        var check = Task.Run(owner.OperationalCheck);
        Thread.Sleep(200);
        var sw = Stopwatch.StartNew();
        thread.Run(() => { });                                     // the owner thread still answers during the check
        Assert.True(sw.ElapsedMilliseconds < 500, $"owner thread blocked {sw.ElapsedMilliseconds} ms");
        Assert.True(await check.WaitAsync(TimeSpan.FromSeconds(10)));
        owner.RemoveAndRelease();
        thread.Run(() => mutex!.Dispose());
    }

    // ---- review fix 4: a broken test configuration never stops a normal start

    [Fact]
    public void Normal_start_ignores_an_incomplete_or_invalid_test_configuration()
    {
        Assert.Null(HelperConfig.ForNormalStart(_ => null, out var none));
        Assert.Null(none);
        Assert.Null(HelperConfig.ForNormalStart(Env(@"C:\x", null), out var partial));
        Assert.Contains("Test configuration ignored", partial);
        Assert.Null(HelperConfig.ForNormalStart(Env(null, "0123456789ab"), out var partial2));
        Assert.NotNull(partial2);
        Assert.Null(HelperConfig.ForNormalStart(Env(Path.GetTempPath(), "BAD"), out var invalid));
        Assert.NotNull(invalid);
        Assert.Throws<InvalidOperationException>(() => HelperConfig.From(Env(@"C:\x", null)));   // --handoff still refuses
    }

    // ---- production: only the installed helper (LocalAppData\PPTK\versions\<v>\PPTKKeys\PPTKKeys.dll) updates itself

    [Theory]
    [InlineData(@"C:\U\Local\PPTK\versions\1.2.3\PPTKKeys\PPTKKeys.dll", true)]
    [InlineData(@"C:\U\Local\PPTK\VERSIONS\10.0.1\pptkkeys\pptkkeys.dll", true)]
    [InlineData(@"C:\U\Local\PPTK\PPTKKeys\PPTKKeys.dll", false)]                  // manual (legacy) installation
    [InlineData(@"C:\GitHub\PPTK\src\PPTKKeys\windows\bin\PPTKKeys.dll", false)]  // development run
    [InlineData(@"C:\U\Local\PPTK\versions\1.2\PPTKKeys\PPTKKeys.dll", false)]
    [InlineData(@"C:\U\Local\PPTK\versions\1.2.3\Other\PPTKKeys.dll", false)]
    [InlineData(@"C:\U\Local\PPTK\versions\1.2.3\PPTKKeys\sub\PPTKKeys.dll", false)]
    [InlineData(@"C:\U\Local\PPTK\versions\..\versions\1.2.3\x\..\PPTKKeys\Other.dll", false)]
    [InlineData(@"C:\Other\PPTK\versions\1.2.3\PPTKKeys\PPTKKeys.dll", false)]
    [InlineData(null, false)]
    public void Production_configuration_only_for_the_installed_helper(string? dll, bool expected)
    {
        var c = HelperConfig.ForProduction(dll, @"C:\U\Roaming", @"C:\U\Local");
        Assert.Equal(expected, c is not null);
        if (c is null) return;
        Assert.False(c.IsTest);
        Assert.Equal(@"C:\U\Local\PPTK", c.Paths.PptkDir);
        Assert.Equal(@"C:\U\Roaming\Microsoft\AddIns\PPTK.ppam", c.Paths.AddinFile);
        Assert.Equal(HandoffNames.ProductionMutex, c.Names.Mutex);
        Assert.True(c.Live.IsProduction);
        Assert.Equal("PPTKKeys Startup", c.Live.TaskName);
    }

    // ---- review fix 5: how the new helper is started is resolved, not assumed

    [Fact]
    public void Dotnet_host_is_resolved_whatever_started_the_old_helper()
    {
        const string local = @"C:\U\AppData\Local", pf = @"C:\Program Files";
        string userDotnet = local + @"\Microsoft\dotnet\dotnet.exe", machineDotnet = pf + @"\dotnet\dotnet.exe";
        var present = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { @"D:\sdk\dotnet.exe", userDotnet, machineDotnet };
        Func<string, bool> exists = present.Contains;
        Assert.Equal(@"D:\sdk\dotnet.exe", HandoffLaunch.ResolveDotnet(@"D:\sdk\dotnet.exe", exists, local, pf));        // dotnet exec
        Assert.Equal(userDotnet, HandoffLaunch.ResolveDotnet(@"C:\x\PPTKKeys.exe", exists, local, pf));                  // apphost
        Assert.Equal(userDotnet, HandoffLaunch.ResolveDotnet(null, exists, local, pf));
        present.Remove(userDotnet);
        Assert.Equal(machineDotnet, HandoffLaunch.ResolveDotnet(@"C:\x\PPTKKeys.exe", exists, local, pf));
        present.Clear();
        Assert.Null(HandoffLaunch.ResolveDotnet(@"C:\x\PPTKKeys.exe", exists, local, pf));                              // no host: no handoff
        Assert.Null(HandoffLaunch.ResolveDotnet(@"D:\sdk\dotnet.exe", exists, local, pf));                              // stale path

        // Universal versions: the new helper starts on its own private runtime, whatever host the old one used.
        string newDll = local + @"\PPTK\versions\1.1.0\PPTKKeys\PPTKKeys.dll", own = local + @"\PPTK\versions\1.1.0\runtime\dotnet.exe";
        present.Add(own); present.Add(userDotnet);
        Assert.Equal(own, HandoffLaunch.ResolveDotnet(userDotnet, exists, local, pf, newDll));
        present.Remove(own);
        Assert.Equal(userDotnet, HandoffLaunch.ResolveDotnet(userDotnet, exists, local, pf, newDll));                 // no private runtime: as before
    }

    [Fact]
    public void New_helper_target_must_be_a_PPTKKeys_dll_under_versions()
    {
        var paths = new PendingUpdatePaths(@"C:\R\AppData", @"C:\R\LocalAppData");
        string good = @"C:\R\LocalAppData\PPTK\versions\1.0.1\PPTKKeys\PPTKKeys.dll";
        Func<string, bool> exists = p => p.Equals(good, StringComparison.OrdinalIgnoreCase);
        Assert.Null(HandoffLaunch.CheckNewHelper(good, paths, exists));
        Assert.NotNull(HandoffLaunch.CheckNewHelper(null, paths, exists));
        Assert.NotNull(HandoffLaunch.CheckNewHelper(@"C:\Elsewhere\PPTKKeys.dll", paths, _ => true));
        Assert.NotNull(HandoffLaunch.CheckNewHelper(@"C:\R\LocalAppData\PPTK\versions\1.0.1\PPTKKeys\Other.dll", paths, _ => true));
        Assert.NotNull(HandoffLaunch.CheckNewHelper(@"C:\R\LocalAppData\PPTK\versions\..\..\..\Evil\PPTKKeys.dll", paths, _ => true));
        Assert.NotNull(HandoffLaunch.CheckNewHelper(@"C:\R\LocalAppData\PPTK\versions\1.0.2\PPTKKeys\PPTKKeys.dll", paths, exists));   // missing
    }

    // ---- configuration

    static Func<string, string?> Env(string? root, string? id) => n => n == "PPTK_TEST_ROOT" ? root : n == "PPTK_TEST_ID" ? id : null;

    [Fact]
    public void Without_test_variables_there_is_no_configuration()
    {
        Assert.Null(HelperConfig.From(_ => null));
    }

    [Fact]
    public void Incomplete_or_unsafe_test_configurations_are_refused()
    {
        var root = Path.Combine(Path.GetTempPath(), "PPTK-config-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(root);
        try
        {
            Assert.Throws<InvalidOperationException>(() => HelperConfig.From(Env(root, null)));
            Assert.Throws<InvalidOperationException>(() => HelperConfig.From(Env(null, "0123456789ab")));
            Assert.Throws<InvalidOperationException>(() => HelperConfig.From(Env(root, "NOT-HEX")));
            Assert.Throws<InvalidOperationException>(() => HelperConfig.From(Env(root, "0123456789ab")));   // no PPTK-TEST-ROOT marker
            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string roaming = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            foreach (var real in new[] { local, Path.Combine(local, "PPTK"), Path.Combine(local, "PPTK", "x"), roaming, Path.Combine(roaming, "Microsoft", "AddIns") })
            {
                var ex = Assert.Throws<InvalidOperationException>(() => HelperConfig.From(Env(real, "0123456789ab")));
                Assert.Contains("overlaps the real folder", ex.Message);   // refused before anything is read there
            }
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void Valid_test_configuration_uses_test_names_only()
    {
        var root = Path.Combine(Path.GetTempPath(), "PPTK-config-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(root);
        File.WriteAllText(Path.Combine(root, "PPTK-TEST-ROOT"), "test");
        try
        {
            var c = HelperConfig.From(Env(root, "0123456789ab"))!;
            Assert.Equal(@"Local\PPTKKeys.Test.0123456789ab", c.Names.Mutex);
            Assert.NotEqual(HandoffNames.ProductionMutex, c.Names.Mutex);
            Assert.Equal("PPTK-Test-0123456789ab", c.Live.TaskName);
            Assert.False(c.Live.IsProduction);
            Assert.StartsWith(root, c.Paths.AddinFile);
            Assert.StartsWith(root, c.LogPath);
            Assert.Null(c.RequestedHelperDll());
        }
        finally { Directory.Delete(root, true); }
    }

    // ---- when the handoff starts

    [Fact]
    public void Handoff_starts_once_and_only_for_an_applied_update_that_changed_the_helper()
    {
        var started = new List<string>();
        var t = new HandoffTrigger(r => started.Add(r.HelperDll!));
        t.OnResult(new PendingUpdateResult(PendingUpdateOutcome.Deferred, "PowerPoint is running"));
        t.OnResult(new PendingUpdateResult(PendingUpdateOutcome.Applied, "x", "1.0.1", HelperRestartRequired: false, HelperDll: @"C:\v\1.0.1\PPTKKeys.dll"));
        t.OnResult(new PendingUpdateResult(PendingUpdateOutcome.RolledBack, "x"));
        Assert.Empty(started);
        Assert.False(t.Started);
        t.OnResult(new PendingUpdateResult(PendingUpdateOutcome.Applied, "x", "1.0.1", HelperRestartRequired: true, HelperDll: @"C:\v\1.0.1\PPTKKeys.dll"));
        t.OnResult(new PendingUpdateResult(PendingUpdateOutcome.Applied, "x", "1.0.2", HelperRestartRequired: true, HelperDll: @"C:\v\1.0.2\PPTKKeys.dll"));
        Assert.Equal(new[] { @"C:\v\1.0.1\PPTKKeys.dll" }, started);
        Assert.True(t.Started);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Watcher_reports_the_applied_update_to_the_handoff_trigger(bool keysChanged)
    {
        using var f = new PendingFixture(keysChanged);
        var started = new List<string>();
        var trigger = new HandoffTrigger(r => started.Add(r.HelperDll!));
        var w = new PendingUpdateWatcher(f.Paths, () => f.Engine(), new FakeProbe(), _ => { }, onResult: trigger.OnResult);
        Assert.True(w.Trigger("startup"));
        await w.Current.WaitAsync(TimeSpan.FromSeconds(30));
        Assert.Equal(PendingUpdateOutcome.Applied, w.LastResult!.Outcome);
        Assert.Equal(keysChanged ? new[] { f.Paths.HelperDll(f.Pending) } : [], started);
    }

    // ---- the real helper binary: only arguments that exit before any hook is installed

    static string HelperDll()
    {
        var bin = new DirectoryInfo(AppContext.BaseDirectory.TrimEnd('\\'));
        string config = bin.Parent!.Name, windows = bin.Parent!.Parent!.Parent!.Parent!.FullName;
        var dll = Path.Combine(windows, "bin", config, "net10.0-windows", "PPTKKeys.dll");
        Assert.True(File.Exists(dll), "build src\\PPTKKeys\\windows\\PPTKKeys.csproj first: " + dll);
        return dll;
    }

    static (int Code, string Out, string Err) Helper(Dictionary<string, string>? env, params string[] args)
    {
        var psi = new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\dotnet\dotnet.exe"))
        { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
        psi.ArgumentList.Add("exec"); psi.ArgumentList.Add(HelperDll());
        foreach (var a in args) psi.ArgumentList.Add(a);
        psi.Environment.Remove("PPTK_TEST_ROOT"); psi.Environment.Remove("PPTK_TEST_ID");
        foreach (var (k, v) in env ?? []) psi.Environment[k] = v;
        using var p = Process.Start(psi)!;
        var err = p.StandardError.ReadToEndAsync();
        if (!p.WaitForExit(30_000)) { p.Kill(); p.WaitForExit(); Assert.Fail("the helper did not exit (killed)"); }
        string o = p.StandardOutput.ReadToEnd();
        return (p.ExitCode, o.Trim(), err.Result.Trim());
    }

    static bool InstalledHelperRunning()
    {
        bool running = Mutex.TryOpenExisting(HandoffNames.ProductionMutex, out var m);
        m?.Dispose();
        return running;
    }

    [Fact]
    public void Helper_informational_arguments_behave_as_before()
    {
        var v = Helper(null, "--version");
        Assert.Equal(0, v.Code); Assert.Equal("1.0.0", v.Out);
        var h = Helper(null, "--help");
        Assert.Equal(0, h.Code); Assert.Contains("Usage: PPTKKeys", h.Out);
        Assert.Equal(1, Helper(null, "--bogus").Code);
    }

    [Fact]
    public void Handoff_mode_is_refused_without_a_test_configuration()
    {
        var r = Helper(null, "--handoff", "1234", "0123456789abcdef");
        Assert.Equal(1, r.Code);
        Assert.Contains("production update activation is disabled", r.Err);
        Assert.Equal(1, Helper(null, "--handoff", "x", "y").Code);
    }

    [Fact]
    public void A_partial_test_configuration_does_not_stop_the_normal_start_path()
    {
        // Fix 4 at the binary level: with only PPTK_TEST_ROOT set the helper warns and continues on the normal path,
        // where the installed helper's mutex makes it exit at once ("already running"), before any window or hook.
        Assert.True(InstalledHelperRunning(), "this check needs the installed helper running; it is skipped otherwise");
        var r = Helper(new Dictionary<string, string> { ["PPTK_TEST_ROOT"] = Path.GetTempPath() });
        Assert.Equal(1, r.Code);
        Assert.Contains("Test configuration ignored", r.Err);
        Assert.DoesNotContain("refused", r.Err);
        Assert.True(InstalledHelperRunning());
    }

    [Fact]
    public void A_test_helper_refuses_to_start_next_to_the_installed_helper()
    {
        // Only meaningful (and only safe) while the installed helper runs: then nothing below can reach a hook.
        Assert.True(InstalledHelperRunning(), "this check needs the installed helper running; it is skipped otherwise");
        var root = Path.Combine(Path.GetTempPath(), "PPTK-config-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(root);
        File.WriteAllText(Path.Combine(root, "PPTK-TEST-ROOT"), "test");
        try
        {
            var env = new Dictionary<string, string> { ["PPTK_TEST_ROOT"] = root, ["PPTK_TEST_ID"] = "0123456789ab" };
            var normal = Helper(env);
            Assert.Equal(1, normal.Code);
            Assert.Contains("installed PPTKKeys helper is running", normal.Err);
            var handoff = Helper(env, "--handoff", "1234", "0123456789abcdef");
            Assert.Equal(1, handoff.Code);
            Assert.Contains("installed PPTKKeys helper is running", handoff.Err);
            Assert.False(Directory.Exists(Path.Combine(root, "LocalAppData")));   // refused before writing anything
        }
        finally { Directory.Delete(root, true); }
        Assert.True(InstalledHelperRunning());
    }

    [Fact]
    public void Normal_start_still_exits_when_the_helper_is_already_running()
    {
        // The unchanged normal path: the installed helper owns the mutex, so this start must exit at once
        // (before any window or hook). Only run while the installed helper is running.
        Assert.True(InstalledHelperRunning(), "this check needs the installed helper running; it is skipped otherwise");
        var r = Helper(null);
        Assert.Equal(1, r.Code);
        Assert.True(InstalledHelperRunning());
    }
}
