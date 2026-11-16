using System.Runtime.Versioning;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace PPTKKeys;

// Phase 2B-3c-1: the pieces the helper uses for pending updates and the handoff, kept free of Win32 so the
// tests can drive them with a fake hook. Active only with a test configuration (HelperConfig); production
// activation stays off (DeploymentFlags).

/// <summary>The helper mutex and keyboard hook, owned together. Every operation runs on the owner thread (the
/// helper's message thread: the mutex is thread-affine) through <c>onOwnerThread</c>, which runs the work there
/// and returns when it is done.</summary>
public sealed class HelperHookOwner : IHookOwner
{
    readonly Mutex mutex;
    readonly Func<IntPtr> install;
    readonly Action<IntPtr> uninstall;
    readonly Func<long> timerTicks;
    readonly Action<Action> onOwnerThread;
    readonly TimeSpan timerWait;
    IntPtr hook;
    bool held;

    /// <param name="mutexHeld">True for a normal start: Main already owns the mutex (whatever happens to the hook,
    /// Main's finally then releases it, as before). False for a helper started to take over (--handoff).</param>
    /// <param name="install">Installs the hook, returns its handle (zero on failure). Runs on the owner thread.</param>
    /// <param name="timerTicks">A counter the helper's 1-second timer increases (operational check).</param>
    /// <param name="timerWait">How long the operational check waits for a tick. It polls from the handoff thread and
    /// never blocks the message thread; 5 s leaves room for a slow PowerPoint refresh on that thread.</param>
    public HelperHookOwner(Mutex mutex, bool mutexHeld, Func<IntPtr> install, Action<IntPtr> uninstall, Func<long> timerTicks, Action<Action> onOwnerThread, TimeSpan? timerWait = null)
    {
        this.mutex = mutex; held = mutexHeld; this.install = install; this.uninstall = uninstall; this.timerTicks = timerTicks; this.onOwnerThread = onOwnerThread;
        this.timerWait = timerWait ?? TimeSpan.FromSeconds(5);
    }

    public bool Held { get { bool r = false; onOwnerThread(() => r = held); return r; } }
    public IntPtr Hook { get { IntPtr r = IntPtr.Zero; onOwnerThread(() => r = hook); return r; } }

    /// <summary>Normal start: the hook the helper installed itself (unchanged startup path).</summary>
    public void Adopt(IntPtr installedHook) => onOwnerThread(() => hook = installedHook);

    public bool AcquireAndInstall(TimeSpan mutexTimeout)
    {
        bool ok = false;
        onOwnerThread(() =>
        {
            if (!held)
            {
                try { if (!mutex.WaitOne(mutexTimeout)) return; }
                catch (AbandonedMutexException) { }   // the previous owner died: the mutex is ours
                held = true;
            }
            if (hook == IntPtr.Zero)
            {
                hook = install();
                if (hook == IntPtr.Zero) { held = false; mutex.ReleaseMutex(); return; }
            }
            ok = true;
        });
        return ok;
    }

    public void RemoveAndRelease() => onOwnerThread(() =>
    {
        if (hook != IntPtr.Zero) { uninstall(hook); hook = IntPtr.Zero; }
        if (held) { held = false; mutex.ReleaseMutex(); }
    });

    /// <summary>The hook is installed, the owner thread answers (this call goes through it) and the timer ticks.</summary>
    public bool OperationalCheck()
    {
        bool hooked = false;
        onOwnerThread(() => hooked = hook != IntPtr.Zero);
        if (!hooked) return false;
        long start = timerTicks();
        var until = DateTime.UtcNow + timerWait;
        while (DateTime.UtcNow < until) { if (timerTicks() != start) return true; Thread.Sleep(50); }
        return false;
    }
}

/// <summary>Test configuration from PPTK_TEST_ROOT + PPTK_TEST_ID. Absent: null, and the helper runs exactly as
/// before (no watcher, no handoff, no task or registry access).</summary>
[SupportedOSPlatform("windows")]
public sealed class HelperConfig
{
    static readonly Regex Id = new("^[0-9a-f]{12}$", RegexOptions.CultureInvariant);

    HelperConfig(string root, string id)
    {
        Root = root; TestId = id; IsTest = true;
        Paths = new PendingUpdatePaths(Path.Combine(root, "AppData"), Path.Combine(root, "LocalAppData"));
        Live = LiveDeploymentConfig.ForTest("PPTK-Test-" + id, @"Software\PPTK-Test-" + id);
        Names = HandoffNames.ForTest(@"Local\PPTKKeys.Test." + id);
    }

    HelperConfig(PendingUpdatePaths paths)
    {
        Root = ""; TestId = ""; IsTest = false;
        Paths = paths;
        Live = LiveDeploymentConfig.Production();
        Names = HandoffNames.Production();
    }

    public string Root { get; }
    public string TestId { get; }
    public bool IsTest { get; }
    public PendingUpdatePaths Paths { get; }
    public LiveDeploymentConfig Live { get; }
    public HandoffNames Names { get; }
    public string LogPath => Path.Combine(Paths.PptkDir, "logs", $"helper-{DateTime.UtcNow:yyyyMMdd}.log");

    public static HelperConfig? FromEnvironment() => From(Environment.GetEnvironmentVariable);

    /// <summary>For a normal start: an incomplete or unsafe test configuration is ignored (with a warning) so it can
    /// never stop the helper from starting normally. (--handoff still requires a valid one: From throws.)</summary>
    public static HelperConfig? ForNormalStart(Func<string, string?> env, out string? warning)
    {
        warning = null;
        try { return From(env); }
        catch (InvalidOperationException ex) { warning = "Test configuration ignored: " + ex.Message; return null; }
    }

    /// <summary>Throws (the helper then refuses to start) when the test configuration is incomplete or unsafe.</summary>
    public static HelperConfig? From(Func<string, string?> env)
    {
        string? root = env("PPTK_TEST_ROOT"), id = env("PPTK_TEST_ID");
        if (root is null && id is null) return null;
        if (string.IsNullOrEmpty(root) || string.IsNullOrEmpty(id)) throw new InvalidOperationException("PPTK_TEST_ROOT and PPTK_TEST_ID must be set together");
        if (!Id.IsMatch(id)) throw new InvalidOperationException("PPTK_TEST_ID must be 12 lowercase hex characters");
        root = Path.GetFullPath(root).TrimEnd('\\');
        foreach (var live in new[] { Environment.SpecialFolder.ApplicationData, Environment.SpecialFolder.LocalApplicationData })
        {
            string real = Path.GetFullPath(Environment.GetFolderPath(live)).TrimEnd('\\');
            // The test root may live under %TEMP% (inside LocalAppData) but never at or above the PPTK folders.
            foreach (var pptk in new[] { Path.Combine(real, "PPTK"), Path.Combine(real, "Microsoft", "AddIns") })
                if (Under(root, pptk) || Under(pptk, root)) throw new InvalidOperationException($"test root {root} overlaps the real folder {pptk}");
        }
        if (!File.Exists(Path.Combine(root, "PPTK-TEST-ROOT"))) throw new InvalidOperationException($"{root} is not a PPTK test root (PPTK-TEST-ROOT missing)");
        return new HelperConfig(root, id);
    }

    /// <summary>The installed helper: this DLL is LocalAppData\PPTK\versions\&lt;MAJOR.MINOR.PATCH&gt;\PPTKKeys\PPTKKeys.dll of
    /// the signed-in user (the layout install.ps1 writes). Any other location (a manual or legacy installation, a
    /// development run) gets null: the helper then runs without pending updates, notice or handoff, as before.</summary>
    public static HelperConfig? ForProduction(string? dllPath, string appData, string localAppData)
    {
        if (!DeploymentFlags.ProductionUpdatesEnabled || string.IsNullOrEmpty(dllPath)) return null;
        string full = Path.GetFullPath(dllPath);
        string versions = Path.GetFullPath(Path.Combine(localAppData, "PPTK", "versions")).TrimEnd('\\') + "\\";
        if (!full.StartsWith(versions, StringComparison.OrdinalIgnoreCase)) return null;
        string[] rest = full[versions.Length..].Split('\\');
        if (rest.Length != 3 || !Regex.IsMatch(rest[0], @"^\d+\.\d+\.\d+$") || !rest[1].Equals("PPTKKeys", StringComparison.OrdinalIgnoreCase)
            || !rest[2].Equals("PPTKKeys.dll", StringComparison.OrdinalIgnoreCase)) return null;
        return new HelperConfig(new PendingUpdatePaths(appData, localAppData));
    }

    static bool Under(string path, string dir) => path.Equals(dir, StringComparison.OrdinalIgnoreCase) || path.StartsWith(dir + "\\", StringComparison.OrdinalIgnoreCase);

    /// <summary>The handoff target from state.json: the DLL the update asked to activate, or null.</summary>
    public string? RequestedHelperDll()
    {
        try { return (string?)JsonNode.Parse(File.ReadAllText(Paths.StatePath))?["helperActivation"]?["dll"]; }
        catch (Exception e) when (e is IOException or System.Text.Json.JsonException) { return null; }
    }
}

/// <summary>The helper's message thread has ended (window closed) before queued hook/mutex work could run there.
/// Main's finally then removes the hook and releases the mutex on that thread itself.</summary>
public sealed class HelperThreadGoneException() : Exception("the helper thread is no longer running");

/// <summary>Runs one side of the handoff and turns every outcome (and any exception, e.g. the helper window closing
/// mid-handoff) into a logged result: it never throws, so a background handoff thread can never crash the helper.</summary>
[SupportedOSPlatform("windows")]
public sealed class HandoffCoordinator(HandoffNames names, IHookOwner owner, HandoffTimeouts timeouts, Action<string> log,
    Action<string, string?> setStatus, Action closeHelper, int lateRecoveryAttempts = 10)
{
    /// <summary>Old side. Null when the handoff was cut short (e.g. this helper is closing).</summary>
    /// <param name="startNew">Starts the new helper with the given nonce on its command line.</param>
    public OldHandoffOutcome? HandOver(Func<string, System.Diagnostics.Process?> startNew)
    {
        try
        {
            string nonce = HandoffNames.NewNonce();
            var outcome = HelperHandoff.RunOld(names, nonce, () => startNew(nonce), owner, timeouts, log);
            log("handoff (old): " + outcome);
            if (outcome == OldHandoffOutcome.RecoveryFailed)
            {
                // Something else holds the mutex. Keep trying to take the hook back rather than leaving no helper.
                for (int i = 1; i <= lateRecoveryAttempts; i++)
                {
                    if (owner.AcquireAndInstall(timeouts.Recover)) { log($"handoff (old): recovered late (attempt {i})"); outcome = OldHandoffOutcome.Recovered; break; }
                }
            }
            if (outcome != OldHandoffOutcome.Completed) setStatus("failed", "handoff: " + outcome);
            if (outcome is OldHandoffOutcome.Completed or OldHandoffOutcome.RecoveryFailed) closeHelper();
            return outcome;
        }
        catch (Exception ex)
        {
            log($"handoff (old): stopped ({ex.GetType().Name}: {ex.Message})");
            try { setStatus("failed", "handoff stopped: " + ex.Message); } catch { }
            return null;
        }
    }

    /// <summary>New side. Null when the handoff was cut short; the helper then closes.</summary>
    public NewHandoffOutcome? TakeOver(int oldPid, string nonce, Func<bool> verify, Action onActive)
    {
        try
        {
            var outcome = HelperHandoff.RunNew(names, oldPid, nonce, verify, owner, () => setStatus("active", null), timeouts, log);
            log("handoff (new): " + outcome);
            if (outcome == NewHandoffOutcome.Active) onActive(); else closeHelper();
            return outcome;
        }
        catch (Exception ex)
        {
            log($"handoff (new): stopped ({ex.GetType().Name}: {ex.Message})");
            try { closeHelper(); } catch { }
            return null;
        }
    }
}

/// <summary>How the old helper starts the new one: "&lt;dotnet.exe&gt; exec &lt;new dll&gt; --handoff &lt;pid&gt; &lt;nonce&gt;". The old
/// helper may have been started some other way (apphost PPTKKeys.exe, run.cmd), so the host is resolved, not assumed.</summary>
public static class HandoffLaunch
{
    /// <summary>This process's host when it is dotnet.exe; otherwise the first dotnet.exe in the installer's order
    /// (per-user, then machine-wide). Null when none exists: the handoff is then not attempted.</summary>
    public static string? ResolveDotnet(string? processPath, Func<string, bool> exists, string localAppData, string programFiles, string? newDll = null)
    {
        // A universal version runs on its own private runtime: versions\<v>\runtime\dotnet.exe next to versions\<v>\PPTKKeys.
        if (newDll is not null && Path.GetDirectoryName(Path.GetDirectoryName(newDll)) is { } versionDir)
        {
            string own = Path.Combine(versionDir, "runtime", "dotnet.exe");
            if (exists(own)) return own;
        }
        if (processPath is not null && Path.GetFileName(processPath).Equals("dotnet.exe", StringComparison.OrdinalIgnoreCase) && exists(processPath)) return processPath;
        foreach (var candidate in new[] { Path.Combine(localAppData, "Microsoft", "dotnet", "dotnet.exe"), Path.Combine(programFiles, "dotnet", "dotnet.exe") })
            if (exists(candidate)) return candidate;
        return null;
    }

    /// <summary>The new helper DLL must be the one the update asked for, exist, and live under versions\.</summary>
    public static string? CheckNewHelper(string? dll, PendingUpdatePaths paths, Func<string, bool> exists)
    {
        if (dll is null) return "no new helper was named";
        string full = Path.GetFullPath(dll), versions = Path.GetFullPath(Path.Combine(paths.PptkDir, "versions")) + "\\";
        if (!full.StartsWith(versions, StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(full).Equals("PPTKKeys.dll", StringComparison.OrdinalIgnoreCase))
            return $"the new helper '{dll}' is not a PPTKKeys.dll under {versions}";
        return exists(full) ? null : $"the new helper '{dll}' does not exist";
    }
}

/// <summary>Starts the handoff once, and only for an applied update that changed the helper files.</summary>
public sealed class HandoffTrigger(Action<PendingUpdateResult> start)
{
    int started;

    public bool Started => Volatile.Read(ref started) == 1;

    public void OnResult(PendingUpdateResult result)
    {
        if (result.Outcome != PendingUpdateOutcome.Applied || !result.HelperRestartRequired || result.HelperDll is null) return;
        if (Interlocked.CompareExchange(ref started, 1, 0) != 0) return;
        start(result);
    }
}

/// <summary>Append-only helper log (test configuration only). Never throws.</summary>
public sealed class HelperLog(string path)
{
    readonly object gate = new();
    public void Write(string message)
    {
        try
        {
            lock (gate)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                File.AppendAllText(path, $"{DateTime.UtcNow:o} [{Environment.ProcessId}] {message}\r\n");
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }
}
