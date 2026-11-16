using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace PPTKKeys;

// Phase 2A: applies an add-in update that src\Deploy\install.ps1 staged while PowerPoint was open.
// Not called by the running helper yet (Phase 2B). No UI, no keyboard, no network, no process control.
//
// Files and formats are the ones install.ps1 writes (LocalAppData\PPTK):
//   state.json        {schemaVersion 1, installed, previous, pending{version, file, sha256, keysChanged, ...}}
//   state.lock        opened with FileShare.None by every PPTK operation (install, uninstall, this engine)
//   transaction.json  undo journal; the same fields as install.ps1, so either side can undo the other's run
//   versions\<v>\     {version.json, PPTK.ppam, PPTKKeys\...}
//   pending\PPTK-<v>.ppam

/// <summary>Scheduled task, registry and process access. Tests use a simulation; Phase 2B adds the live one.</summary>
public interface IDeploymentEnvironment
{
    bool IsPowerPointRunning();
    /// <summary>The journal's "tasksFormat" for this environment: "sim-json" (test roots) or "task-xml" (live,
    /// not implemented yet). A journal written in another format is never restored here (src\Deploy\JOURNAL.md).</summary>
    string TasksFormat { get; }
    /// <summary>Snapshot of the scheduled tasks in the journal's "tasksBefore" format; null when there are none.</summary>
    string? ExportTasks();
    /// <summary>Restores a snapshot; null removes only the PPTKKeys task(s) and keeps every other task.</summary>
    void ImportTasks(string? snapshot);
    /// <summary>The DLL the "PPTKKeys Startup" task runs (dotnet exec "&lt;dll&gt;"); null without that task.</summary>
    string? GetHelperTaskDll();
    /// <summary>The program the task starts (the dotnet.exe host); null without that task.</summary>
    string? GetHelperTaskHost();
    /// <summary>Points the existing task at another DLL and, when <paramref name="host"/> is given, another dotnet.exe
    /// (the private runtime of that version). Does not stop or start anything.</summary>
    void SetHelperTaskDll(string dll, string? host = null);
    /// <summary>HKCU key values as a JSON object; null when the key does not exist.</summary>
    string? ExportRegistryKey(string key);
    void ImportRegistryKey(string key, string? json);
}

public sealed class PendingUpdatePaths
{
    public PendingUpdatePaths(string appData, string localAppData)
    {
        PptkDir = Path.Combine(localAppData, "PPTK");
        AddinFile = Path.Combine(appData, "Microsoft", "AddIns", "PPTK.ppam");
    }

    public string PptkDir { get; }
    public string AddinFile { get; }
    public string StatePath => Path.Combine(PptkDir, "state.json");
    public string LockPath => Path.Combine(PptkDir, "state.lock");
    public string JournalPath => Path.Combine(PptkDir, "transaction.json");
    public string TxnDir => Path.Combine(PptkDir, "txn");
    public string VersionDir(string version) => Path.Combine(PptkDir, "versions", version);
    public string HelperDll(string version) => Path.Combine(VersionDir(version), "PPTKKeys", "PPTKKeys.dll");
    /// <summary>The private .NET host of a universal version (install.ps1 copies the runtime of this device's
    /// architecture into versions\&lt;v&gt;\runtime\).</summary>
    public string RuntimeHost(string version) => Path.Combine(VersionDir(version), "runtime", "dotnet.exe");
}

public enum PendingUpdateOutcome
{
    NothingPending,
    Applied,
    /// <summary>Not now (PowerPoint running, add-in in use, another operation running). Nothing changed.</summary>
    Deferred,
    /// <summary>The state or the staged files are not valid. Nothing changed except state.lastError.</summary>
    Rejected,
    /// <summary>A step failed; everything was restored to the installed version.</summary>
    RolledBack,
    /// <summary>A step failed and the restore failed too; transaction.json is kept for the next run.</summary>
    RestoreFailed,
    /// <summary>Test hook only: the run stopped without clean-up, as if the process had been killed.</summary>
    Interrupted,
}

/// <param name="HelperRestartRequired">The helper files changed: the task already runs the new DLL from the
/// next start, and the running (old) helper must hand over (Phase 2B). The engine never restarts it.</param>
public sealed record PendingUpdateResult(PendingUpdateOutcome Outcome, string Detail, string? Version = null, bool HelperRestartRequired = false, string? HelperDll = null);

/// <summary>Result of PendingAddinUpdate.Inspect: is there a pending update, is it verified, and was the notice shown.</summary>
public sealed record PendingInspection(bool HasPending, bool Valid, string? Version, bool Notified, string? Problem);

public sealed class PendingUpdateTestHooks
{
    /// <summary>Called after each step with the step name and the file it concerns. May throw or change files.</summary>
    public Action<string, string?>? AfterStep { get; init; }
    /// <summary>Return true to stop right after that step without any clean-up (simulated crash).</summary>
    public Func<string, bool>? AbandonAfter { get; init; }
    /// <summary>Before each replace attempt (operation, temporary source, target): return an exception to make that
    /// attempt fail. May also change files first (e.g. complete the replace itself) to simulate partial success.</summary>
    public Func<string, string, string, Exception?>? FailOperation { get; init; }
}

public sealed class PendingAddinUpdate
{
    public const string Operation = "apply-pending";
    public const string PptkKey = @"Software\PPTK";
    public const string AddinKey = @"Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK";
    static readonly string[] KeysFiles = ["PPTKKeys.dll", "PPTKKeys.deps.json", "PPTKKeys.runtimeconfig.json"];
    static readonly Regex SemVer = new(@"^\d+\.\d+\.\d+$", RegexOptions.CultureInvariant);
    static readonly Regex Sha256Hex = new("^[0-9a-f]{64}$", RegexOptions.CultureInvariant);
    static readonly UTF8Encoding Utf8 = new(false);

    readonly PendingUpdatePaths paths;
    readonly IDeploymentEnvironment env;
    readonly int lockedRetries;
    readonly Action<TimeSpan> wait;
    readonly PendingUpdateTestHooks? hooks;

    /// <param name="lockedRetries">Extra attempts while the add-in file is still in use (PowerPoint closing).</param>
    /// <param name="wait">Delay between attempts (tests pass a no-op).</param>
    public PendingAddinUpdate(PendingUpdatePaths paths, IDeploymentEnvironment env, int lockedRetries = 5, Action<TimeSpan>? wait = null, PendingUpdateTestHooks? hooks = null)
    {
        this.paths = paths;
        this.env = env;
        this.lockedRetries = Math.Max(0, lockedRetries);
        this.wait = wait ?? Thread.Sleep;
        this.hooks = hooks;
    }

    public List<string> Log { get; } = [];

    sealed class AbandonedException : Exception { }
    sealed class RejectException(string message) : Exception(message) { }

    /// <summary>Read-only check of the pending update (for the update notice): the same validation TryApply uses,
    /// but no lock, no journal, no writes (not even lastError), and it does not matter whether PowerPoint runs.</summary>
    public PendingInspection Inspect()
    {
        if (File.Exists(paths.JournalPath)) return new(false, false, null, false, "a PPTK operation is in progress");
        JsonObject state;
        try { state = JsonNode.Parse(File.ReadAllText(paths.StatePath)) as JsonObject ?? throw new JsonException("not an object"); }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException) { return new(false, false, null, false, "state.json is missing or corrupt"); }
        if (state["pending"] is not JsonObject pending) return new(false, false, null, false, "no pending update");
        string? version = Str(pending, "version");
        bool notified = pending["notified"] is JsonValue n && n.GetValueKind() == JsonValueKind.True;
        try
        {
            var plan = Validate(state);
            return new(true, true, plan.NewVersion, notified, null);
        }
        catch (RejectException ex) { return new(true, false, version, notified, ex.Message); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return new(true, false, version, notified, ex.Message); }
    }

    public PendingUpdateResult TryApply()
    {
        if (!File.Exists(paths.StatePath) && !File.Exists(paths.JournalPath)) return new(PendingUpdateOutcome.NothingPending, "no managed PPTK installation");
        FileStream? held;
        try { held = new FileStream(paths.LockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return new(PendingUpdateOutcome.Deferred, "another PPTK operation holds state.lock"); }
        try { return ApplyLocked(); }
        catch (AbandonedException) { return new(PendingUpdateOutcome.Interrupted, "test hook: stopped without clean-up"); }
        finally { held.Dispose(); }
    }

    PendingUpdateResult ApplyLocked()
    {
        // An interrupted run: undo our own; an installer run is undone by install.ps1/uninstall.ps1.
        if (File.Exists(paths.JournalPath))
        {
            JsonNode? journal;
            try { journal = JsonNode.Parse(File.ReadAllText(paths.JournalPath)); } catch (JsonException) { journal = null; }
            if (Str(journal, "operation") != Operation) return new(PendingUpdateOutcome.Deferred, "an interrupted installer operation must be completed by install.ps1 first");
            try { Undo(journal!, "an earlier pending update was interrupted"); }
            catch (Exception ex) { return new(PendingUpdateOutcome.RestoreFailed, "restoring the interrupted update failed: " + ex.Message); }
        }
        else TryDeleteDir(paths.TxnDir);   // left by a committed run whose clean-up failed

        JsonObject state;
        try { state = JsonNode.Parse(File.ReadAllText(paths.StatePath)) as JsonObject ?? throw new JsonException("not an object"); }
        catch (Exception ex) when (ex is JsonException or IOException) { return new(PendingUpdateOutcome.Rejected, "state.json is missing or corrupt: " + ex.Message); }
        if (state["pending"] is null) return new(PendingUpdateOutcome.NothingPending, "no pending update");

        if (env.IsPowerPointRunning()) return new(PendingUpdateOutcome.Deferred, "PowerPoint is running");

        Plan plan;
        try { plan = Validate(state); }
        catch (RejectException ex) { SetLastError(state, ex.Message); return new(PendingUpdateOutcome.Rejected, ex.Message); }

        // PowerPoint may still be closing: wait for the add-in to be released (or for a new instance).
        for (int attempt = 0; ; attempt++)
        {
            if (env.IsPowerPointRunning()) return new(PendingUpdateOutcome.Deferred, "PowerPoint is running");
            if (CanOpenExclusively(paths.AddinFile)) break;
            if (attempt >= lockedRetries) return new(PendingUpdateOutcome.Deferred, "the add-in file is in use");
            wait(TimeSpan.FromMilliseconds(500 * (attempt + 1)));
        }

        try
        {
            // The add-in (the only file PowerPoint can hold) is replaced last; after it only small local writes follow.
            SaveJournal(plan.NewVersion);                                   Step("journal", paths.JournalPath);
            env.SetHelperTaskDll(plan.NewDll, plan.NewHost);                Step("task", plan.NewDll);
            if (!SamePath(env.GetHelperTaskDll(), plan.NewDll)) throw new InvalidOperationException("the task does not run the new helper after the change");
            if (plan.NewHost is not null && !SamePath(env.GetHelperTaskHost(), plan.NewHost)) throw new InvalidOperationException("the task does not use the new runtime after the change");
            ReplaceAddin(plan);
            WriteState(NewState(state, plan));                              Step("state", paths.StatePath);
            env.ImportRegistryKey(PptkKey, new JsonObject { ["InstalledVersion"] = plan.NewVersion }.ToJsonString());
            Step("registry", null);
            if (InstalledVersion() != plan.NewVersion) throw new InvalidOperationException("InstalledVersion was not written");
            File.Delete(paths.JournalPath);   // the commit point: from here on nothing is undone
        }
        catch (AbandonedException) { throw; }
        catch (Exception ex)
        {
            if (!File.Exists(paths.JournalPath)) { TryDeleteDir(paths.TxnDir); return new(PendingUpdateOutcome.RolledBack, ex.Message + " (before any change)"); }
            try { Undo(JsonNode.Parse(File.ReadAllText(paths.JournalPath))!, "the update failed"); }
            catch (Exception undo) { return new(PendingUpdateOutcome.RestoreFailed, $"{ex.Message} | RESTORE FAILED: {undo.Message} (transaction.json kept)"); }
            try { SetLastError(JsonNode.Parse(File.ReadAllText(paths.StatePath))!.AsObject(), ex.Message); } catch { }
            return new(PendingUpdateOutcome.RolledBack, ex.Message);
        }

        // Committed. Clean-up is best effort: a failure here never rolls back the applied update
        // (a leftover txn\ without a journal is removed by the next run).
        bool cleaned = TryDeleteDir(paths.TxnDir) & TryDelete(plan.PendingFile);
        try { Directory.Delete(Path.GetDirectoryName(plan.PendingFile)!); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }   // only when empty
        Note($"PPTK {plan.NewVersion} applied (from {plan.OldVersion}); helper restart required: {plan.KeysChanged}" + (cleaned ? "" : "; clean-up left files for the next run"));
        return new(PendingUpdateOutcome.Applied, $"PPTK {plan.NewVersion} applied", plan.NewVersion, plan.KeysChanged, plan.NewDll);
    }

    sealed record Plan(string OldVersion, string NewVersion, string PendingFile, string Sha256, bool KeysChanged, string NewDll, string? NewHost);

    Plan Validate(JsonObject state)
    {
        if (state["schemaVersion"]?.GetValueKind() != JsonValueKind.Number || state["schemaVersion"]!.GetValue<int>() != 1) throw new RejectException("state.json schemaVersion must be 1");
        string? installed = Str(state, "installed");
        if (installed is null || !SemVer.IsMatch(installed)) throw new RejectException("state.json has no valid installed version");
        if (state["pending"] is not JsonObject pending) throw new RejectException("state.json pending entry is not an object");
        string? version = Str(pending, "version");
        if (version is null || !SemVer.IsMatch(version)) throw new RejectException("pending entry has no valid version");
        if (Version.Parse(version) <= Version.Parse(installed)) throw new RejectException($"pending version {version} is not newer than the installed {installed}");
        if (Str(pending, "file") != $@"pending\PPTK-{version}.ppam") throw new RejectException($"unexpected pending file path '{Str(pending, "file")}'");
        string? sha = Str(pending, "sha256");
        if (sha is null || !Sha256Hex.IsMatch(sha)) throw new RejectException("pending entry has no valid sha256");
        if (pending["keysChanged"]?.GetValueKind() is not (JsonValueKind.True or JsonValueKind.False)) throw new RejectException("pending entry has no keysChanged flag");
        bool keysChanged = pending["keysChanged"]!.GetValue<bool>();

        string pendingFile = Path.Combine(paths.PptkDir, "pending", $"PPTK-{version}.ppam");
        if (!File.Exists(pendingFile)) throw new RejectException($@"pending\PPTK-{version}.ppam is missing");
        if (HashFile(pendingFile) != sha) throw new RejectException($@"pending\PPTK-{version}.ppam does not match its SHA-256");

        // The new version: complete and matching its own manifest.
        var newHash = Manifest(version);
        if (newHash.GetValueOrDefault("payload/PPTK.ppam") != sha) throw new RejectException($@"versions\{version}\version.json does not list the pending add-in");
        CheckFiles(version, newHash);

        // The installed version: what the registry, the task and the add-in actually are now.
        if (InstalledVersion() != installed) throw new RejectException($"InstalledVersion does not match state.json ({installed})");
        if (!SamePath(env.GetHelperTaskDll(), paths.HelperDll(installed))) throw new RejectException($@"the PPTKKeys task does not run versions\{installed}");
        var oldHash = Manifest(installed);
        if (!File.Exists(paths.AddinFile) || HashFile(paths.AddinFile) != oldHash.GetValueOrDefault("payload/PPTK.ppam")) throw new RejectException($"the installed add-in is not PPTK {installed}");

        // A universal version brings its own runtime (verified file by file); the task then switches to that host.
        string? newHost = null;
        if (File.Exists(paths.RuntimeHost(version)))
        {
            CheckRuntime(version, newHash, Str(state, "architecture"));
            newHost = paths.RuntimeHost(version);
        }
        else if (IsUnder(env.GetHelperTaskHost(), Path.Combine(paths.PptkDir, "versions")))
            throw new RejectException($@"versions\{version} has no private runtime, but the installed helper uses one");

        bool actual = KeysFiles.Any(f => oldHash.GetValueOrDefault("payload/PPTKKeys/" + f) != newHash["payload/PPTKKeys/" + f])
            || (newHost is not null && !SamePath(env.GetHelperTaskHost(), newHost));
        if (actual != keysChanged) throw new RejectException($"keysChanged is {keysChanged.ToString().ToLowerInvariant()} but the helper files or runtime {(actual ? "differ" : "are identical")}");
        return new Plan(installed, version, pendingFile, sha, keysChanged, paths.HelperDll(version), newHost);
    }

    Dictionary<string, string> Manifest(string version)
    {
        string path = Path.Combine(paths.VersionDir(version), "version.json");
        try
        {
            var m = JsonNode.Parse(File.ReadAllText(path))!;
            if (Str(m, "packageVersion") != version) throw new RejectException($@"versions\{version}\version.json is for another version");
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var f in m["files"]!.AsArray()) map[Str(f, "path")!] = Str(f, "sha256")!;
            return map;
        }
        catch (Exception ex) when (ex is not RejectException) { throw new RejectException($@"versions\{version}\version.json is missing or invalid"); }
    }

    void CheckFiles(string version, Dictionary<string, string> hash)
    {
        var dir = paths.VersionDir(version);
        foreach (var (rel, key) in new[] { ("PPTK.ppam", "payload/PPTK.ppam") }.Concat(KeysFiles.Select(f => (Path.Combine("PPTKKeys", f), "payload/PPTKKeys/" + f))))
        {
            string p = Path.Combine(dir, rel);
            if (!File.Exists(p)) throw new RejectException($@"versions\{version}\{rel} is missing");
            if (!hash.TryGetValue(key, out var want) || HashFile(p) != want) throw new RejectException($@"versions\{version}\{rel} does not match its manifest");
        }
    }

    // Every runtime file the manifest lists for this device's architecture is in versions\<v>\runtime\ with its hash.
    void CheckRuntime(string version, Dictionary<string, string> hash, string? rid)
    {
        if (rid is not ("win-x64" or "win-arm64")) throw new RejectException($"state.json has no valid architecture for the private runtime ('{rid}')");
        string prefix = $"payload/runtime/{rid}/";
        var files = hash.Where(kv => kv.Key.StartsWith(prefix, StringComparison.Ordinal)).ToList();
        if (files.Count == 0) throw new RejectException($@"versions\{version}\version.json lists no {rid} runtime");
        string dir = Path.Combine(paths.VersionDir(version), "runtime");
        foreach (var (key, want) in files)
        {
            string p = Path.Combine(dir, key[prefix.Length..].Replace('/', '\\'));
            if (!File.Exists(p) || HashFile(p) != want) throw new RejectException($@"versions\{version}\runtime\{key[prefix.Length..]} is missing or does not match its manifest");
        }
    }

    static bool IsUnder(string? path, string dir)
    {
        if (string.IsNullOrEmpty(path)) return false;
        string full = Path.GetFullPath(path), d = Path.GetFullPath(dir).TrimEnd('\\') + "\\";
        return full.StartsWith(d, StringComparison.OrdinalIgnoreCase);
    }

    void ReplaceAddin(Plan plan)
    {
        string tmp = Path.Combine(Path.GetDirectoryName(paths.AddinFile)!, "PPTK.ppam.new-" + Guid.NewGuid().ToString("N")[..6]);
        File.Copy(plan.PendingFile, tmp);
        try
        {
            if (HashFile(tmp) != plan.Sha256) throw new InvalidOperationException("the copied add-in does not match the pending SHA-256");
            Step("copied", tmp);
            // Retried only while PowerPoint is still not running: the add-in is never replaced under PowerPoint.
            Retry("replace PPTK.ppam", tmp, paths.AddinFile,
                () => File.Replace(tmp, paths.AddinFile, null),
                done: () => !File.Exists(tmp) && File.Exists(paths.AddinFile) && HashFile(paths.AddinFile) == plan.Sha256,
                guard: () =>
                {
                    if (env.IsPowerPointRunning()) throw new InvalidOperationException("PowerPoint started while the add-in was being replaced");
                    if (!File.Exists(tmp)) throw new InvalidOperationException("the add-in copy disappeared before it could be installed");
                });
        }
        finally { TryDelete(tmp); }
        Step("replaced", paths.AddinFile);
        if (HashFile(paths.AddinFile) != plan.Sha256) throw new InvalidOperationException("the installed add-in does not match the pending SHA-256 after replacing it");
    }

    static JsonObject NewState(JsonObject state, Plan plan)
    {
        var s = state.DeepClone().AsObject();
        s["installed"] = plan.NewVersion;
        s["previous"] = plan.OldVersion;
        s["rolledBackFrom"] = null;
        s["pending"] = null;
        // The task runs the new helper from its next start; Phase 2B hands over the running one.
        s["helperActivation"] = plan.KeysChanged ? new JsonObject { ["version"] = plan.NewVersion, ["dll"] = plan.NewDll, ["requestedUtc"] = DateTime.UtcNow.ToString("o") } : null;
        s["lastError"] = null;
        s["updatedUtc"] = DateTime.UtcNow.ToString("o");
        return s;
    }

    // ------------------------------------------------------------ journal (install.ps1 format)
    void SaveJournal(string version)
    {
        Directory.CreateDirectory(paths.TxnDir);
        string backup = Path.Combine(paths.TxnDir, "PPTK.ppam.before");
        if (File.Exists(backup)) File.SetAttributes(backup, FileAttributes.Normal);
        File.Copy(paths.AddinFile, backup, true);
        File.SetAttributes(backup, FileAttributes.Normal);   // File.Copy keeps ReadOnly; txn\ must stay deletable
        var j = new JsonObject
        {
            ["journalVersion"] = 1,
            ["tasksFormat"] = env.TasksFormat,
            ["operation"] = Operation,
            ["target"] = version,
            ["startedUtc"] = DateTime.UtcNow.ToString("o"),
            ["stateBefore"] = File.ReadAllText(paths.StatePath),
            ["tasksBefore"] = env.ExportTasks(),
            ["pptkKeyBefore"] = env.ExportRegistryKey(PptkKey),
            ["addinKeyBefore"] = env.ExportRegistryKey(AddinKey),
            ["addinBefore"] = backup,
            ["addinBeforeSha"] = HashFile(backup),
            ["createdVersionDir"] = null,
        };
        WriteAtomic(paths.JournalPath, j.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    }

    void Undo(JsonNode j, string why)
    {
        Note($"undo ({why}): restoring the state from before {Str(j, "operation")} {Str(j, "target")}");
        // A journal without these fields was written by the Phase 1 scripts: version 1, simulated tasks.
        if (j["journalVersion"] is JsonValue jv && jv.GetValueKind() == JsonValueKind.Number && jv.GetValue<int>() != 1) throw new InvalidOperationException("unknown journalVersion");
        string format = Str(j, "tasksFormat") ?? "sim-json";
        if (format != env.TasksFormat) throw new InvalidOperationException($"the journal's tasks are in '{format}' format, this environment uses '{env.TasksFormat}'");
        string? before = Str(j, "addinBefore");
        if (before is not null)
        {
            if (HashFile(before) != Str(j, "addinBeforeSha")) throw new InvalidOperationException("the add-in backup in the journal is damaged");
            // Unchanged (e.g. File.Replace failed): do not write over a file PowerPoint may have opened since.
            if (!File.Exists(paths.AddinFile) || HashFile(paths.AddinFile) != Str(j, "addinBeforeSha")) File.Copy(before, paths.AddinFile, true);
            if (HashFile(paths.AddinFile) != Str(j, "addinBeforeSha")) throw new InvalidOperationException("the add-in could not be restored");
        }
        else if (File.Exists(paths.AddinFile)) File.Delete(paths.AddinFile);
        env.ImportTasks(Str(j, "tasksBefore"));
        env.ImportRegistryKey(PptkKey, Str(j, "pptkKeyBefore"));
        env.ImportRegistryKey(AddinKey, Str(j, "addinKeyBefore"));
        string? stateBefore = Str(j, "stateBefore");
        if (stateBefore is not null) WriteAtomic(paths.StatePath, stateBefore); else TryDelete(paths.StatePath);
        CompleteJournal();
    }

    // After a restore: the journal goes first (a kept journal would be undone again), txn\ is best effort.
    void CompleteJournal()
    {
        if (File.Exists(paths.JournalPath)) File.Delete(paths.JournalPath);
        TryDeleteDir(paths.TxnDir);
    }

    // ------------------------------------------------------------ helpers
    void Step(string step, string? file)
    {
        hooks?.AfterStep?.Invoke(step, file);
        if (hooks?.AbandonAfter?.Invoke(step) == true) throw new AbandonedException();
    }

    void Note(string m) => Log.Add(DateTime.UtcNow.ToString("o") + " " + m);

    string? InstalledVersion()
    {
        var json = env.ExportRegistryKey(PptkKey);
        if (json is null) return null;
        try { return Str(JsonNode.Parse(json), "InstalledVersion"); } catch (JsonException) { return null; }
    }

    void SetLastError(JsonObject state, string message)
    {
        state["lastError"] = message;
        state["updatedUtc"] = DateTime.UtcNow.ToString("o");
        try { WriteState(state); } catch (IOException) { }
    }

    void WriteState(JsonObject state) => WriteAtomic(paths.StatePath, state.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));

    static string? Str(JsonNode? node, string name) =>
        node is JsonObject o && o[name] is JsonValue v && v.GetValueKind() == JsonValueKind.String ? v.GetValue<string>() : null;

    static bool SamePath(string? a, string? b) =>
        a is not null && b is not null && string.Equals(Path.GetFullPath(a).TrimEnd('\\'), Path.GetFullPath(b).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase);

    static string HashFile(string path)
    {
        using var s = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Convert.ToHexString(SHA256.HashData(s)).ToLowerInvariant();
    }

    static bool CanOpenExclusively(string path)
    {
        if (!File.Exists(path)) return true;
        try { using var s = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None); return true; }   // read access: a read-only add-in is not "in use"
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return false; }
    }

    void WriteAtomic(string path, string text)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string tmp = path + ".tmp-" + Guid.NewGuid().ToString("N")[..8];
        File.WriteAllText(tmp, text, Utf8);
        try
        {
            Retry("replace " + Path.GetFileName(path), tmp, path,
                () => { if (File.Exists(path)) File.Replace(tmp, path, null); else File.Move(tmp, path); },
                done: () => !File.Exists(tmp) && File.Exists(path) && File.ReadAllText(path) == text,
                guard: () => { if (!File.Exists(tmp)) throw new InvalidOperationException($"the temporary file for {Path.GetFileName(path)} disappeared"); });
        }
        finally { TryDelete(tmp); }
    }

    // Bounded retry for the transient failures seen under load (another process briefly has a just-written
    // file open). Attempts by failure kind; anything else fails at once:
    //   sharing/lock violation, "unable to remove/move" during replace (32, 33, 1175, 1176): 5 attempts
    //   access denied: 3 attempts, none when the target is read-only (a permanent condition)
    // Before a retry, done() says whether the earlier attempt completed after all (then it is not repeated)
    // and guard() throws when retrying is no longer safe.
    internal static int RetryAttempts(Exception ex, string target)
    {
        int code = ex.HResult & 0xFFFF;
        if (ex is IOException && code is 32 or 33 or 1175 or 1176) return 5;
        if (ex is UnauthorizedAccessException || (ex is IOException && code == 5))
            return File.Exists(target) && File.GetAttributes(target).HasFlag(FileAttributes.ReadOnly) ? 1 : 3;
        return 1;
    }

    void Retry(string what, string source, string target, Action op, Func<bool> done, Action guard)
    {
        for (int attempt = 1; ; attempt++)
        {
            if (attempt > 1)
            {
                if (done()) { Note($"{what}: the earlier attempt had completed"); return; }
                guard();
            }
            try
            {
                if (hooks?.FailOperation?.Invoke(what, source, target) is { } simulated) throw simulated;
                op();
                return;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                int max = RetryAttempts(ex, target);
                if (attempt >= max) { if (max > 1) Note($"{what}: attempt {attempt} of {max} failed; giving up"); throw; }
                Note($"{what}: attempt {attempt} failed (0x{ex.HResult:X8} {ex.Message}); retrying");
                wait(TimeSpan.FromMilliseconds(200 * attempt));
            }
        }
    }

    static bool TryDelete(string path) { try { if (File.Exists(path)) File.Delete(path); return true; } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return false; } }
    static bool TryDeleteDir(string path) { try { if (Directory.Exists(path)) Directory.Delete(path, true); return true; } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return false; } }
}
