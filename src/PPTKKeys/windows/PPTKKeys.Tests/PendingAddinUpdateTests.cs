using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

/// <summary>
/// The same simulated environment install.ps1 / detect.ps1 / uninstall.ps1 use in test mode:
/// Root\processes.json, Root\tasks.json, Root\registry.json. Nothing live is touched.
/// </summary>
sealed class SimulatedEnvironment(string root) : IDeploymentEnvironment
{
    static readonly UTF8Encoding Utf8 = new(false);
    string Tasks => Path.Combine(root, "tasks.json");
    string Registry => Path.Combine(root, "registry.json");

    public bool IsPowerPointRunning()
    {
        var p = Path.Combine(root, "processes.json");
        return File.Exists(p) && JsonNode.Parse(File.ReadAllText(p)) is JsonArray a && a.Any(n => string.Equals((string?)n, "POWERPNT", StringComparison.OrdinalIgnoreCase));
    }

    public string TasksFormat => "sim-json";
    public string? ExportTasks() => File.Exists(Tasks) ? File.ReadAllText(Tasks) : null;
    public void ImportTasks(string? snapshot)
    {
        if (snapshot is not null) { File.WriteAllText(Tasks, snapshot, Utf8); return; }
        // null: no PPTKKeys task before; keep every other task.
        var keep = File.Exists(Tasks) && JsonNode.Parse(File.ReadAllText(Tasks)) is JsonArray all
            ? all.OfType<JsonObject>().Where(t => !IsPptk(t)).Select(t => (JsonNode?)t.DeepClone()).ToArray() : [];
        File.WriteAllText(Tasks, new JsonArray(keep).ToJsonString(), Utf8);
    }

    static bool IsPptk(JsonObject t) =>
        ((string?)t["name"] == "PPTKKeys Startup" && (string?)t["path"] == "\\")
        || (t["actions"] as JsonArray ?? []).OfType<JsonObject>().Any(a => $"{(string?)a["execute"]} {(string?)a["arguments"]}".Contains("PPTKKeys.dll", StringComparison.OrdinalIgnoreCase));

    JsonObject? MainTask(JsonArray tasks) => tasks.OfType<JsonObject>().FirstOrDefault(t => (string?)t["name"] == "PPTKKeys Startup" && (string?)t["path"] == "\\");

    public string? GetHelperTaskDll()
    {
        if (!File.Exists(Tasks) || JsonNode.Parse(File.ReadAllText(Tasks)) is not JsonArray tasks || MainTask(tasks) is not { } t) return null;
        var m = Regex.Match((string?)t["actions"]?[0]?["arguments"] ?? "", "^\\s*exec\\s+\"([^\"]+)\"\\s*$");
        return m.Success ? m.Groups[1].Value : null;
    }

    public string? GetHelperTaskHost()
    {
        if (!File.Exists(Tasks) || JsonNode.Parse(File.ReadAllText(Tasks)) is not JsonArray tasks || MainTask(tasks) is not { } t) return null;
        return (string?)t["actions"]?[0]?["execute"];
    }

    public void SetHelperTaskDll(string dll, string? host = null)
    {
        var tasks = JsonNode.Parse(File.ReadAllText(Tasks))!.AsArray();
        var action = MainTask(tasks)!["actions"]![0]!.AsObject();
        action["arguments"] = $"exec \"{dll}\"";
        action["workingDirectory"] = Path.GetDirectoryName(dll);
        if (host is not null) action["execute"] = host;
        File.WriteAllText(Tasks, tasks.ToJsonString(), Utf8);
    }

    JsonObject ReadRegistry() => File.Exists(Registry) ? JsonNode.Parse(File.ReadAllText(Registry))!.AsObject() : new JsonObject();
    public string? ExportRegistryKey(string key) => ReadRegistry()[key]?.ToJsonString();
    public void ImportRegistryKey(string key, string? json)
    {
        var r = ReadRegistry();
        r.Remove(key);
        if (json is not null) r[key] = JsonNode.Parse(json);
        File.WriteAllText(Registry, r.ToJsonString(), Utf8);
    }
}

/// <summary>A managed installation of <c>installed</c> with <c>pending</c> staged, laid out like install.ps1 does.</summary>
sealed class PendingFixture : IDisposable
{
    static readonly UTF8Encoding Utf8 = new(false);
    public string Root { get; }
    public PendingUpdatePaths Paths { get; }
    public SimulatedEnvironment Env { get; }
    public string Installed { get; }
    public string Pending { get; }

    /// <param name="universal">The pending version carries private runtimes (win-x64 and win-arm64 in its manifest,
    /// the win-x64 one staged in versions\&lt;v&gt;\runtime\), as install.ps1 stages a universal package on x64.</param>
    public PendingFixture(bool keysChanged = false, string name = "root", string installed = "1.0.0", string pending = "1.0.1", bool universal = false)
    {
        Universal = universal;
        Root = Path.Combine(Path.GetTempPath(), "PPTK-pending-tests-" + Guid.NewGuid().ToString("N")[..8], name);
        Installed = installed; Pending = pending;
        Directory.CreateDirectory(Root);
        File.WriteAllText(Path.Combine(Root, "PPTK-TEST-ROOT"), "isolated", Utf8);
        Paths = new PendingUpdatePaths(Path.Combine(Root, "AppData"), Path.Combine(Root, "LocalAppData"));
        Env = new SimulatedEnvironment(Root);
        Directory.CreateDirectory(Path.GetDirectoryName(Paths.AddinFile)!);

        WriteVersion(installed, "add-in " + installed, "helper " + installed);
        WriteVersion(pending, "add-in " + pending, "helper " + (keysChanged ? pending : installed));
        File.Copy(Path.Combine(Paths.VersionDir(installed), "PPTK.ppam"), Paths.AddinFile);
        Directory.CreateDirectory(Path.Combine(Paths.PptkDir, "pending"));
        File.Copy(Path.Combine(Paths.VersionDir(pending), "PPTK.ppam"), PendingFile);

        var dotnet = Path.Combine(Root, @"LocalAppData\Microsoft\dotnet\dotnet.exe");
        var dll = Paths.HelperDll(installed);
        File.WriteAllText(Path.Combine(Root, "tasks.json"), new JsonArray(new JsonObject
        {
            ["name"] = "PPTKKeys Startup", ["path"] = "\\", ["enabled"] = true, ["userId"] = @"TESTDOM\tester", ["logonType"] = "Interactive", ["runLevel"] = "Limited",
            ["triggers"] = new JsonArray(new JsonObject { ["type"] = "Logon", ["userId"] = @"TESTDOM\tester", ["enabled"] = true }),
            ["actions"] = new JsonArray(new JsonObject { ["execute"] = dotnet, ["arguments"] = $"exec \"{dll}\"", ["workingDirectory"] = Path.GetDirectoryName(dll) }),
            ["settings"] = new JsonObject { ["multipleInstances"] = "IgnoreNew", ["executionTimeLimit"] = "PT0S" },
        }).ToJsonString(), Utf8);
        File.WriteAllText(Path.Combine(Root, "registry.json"), new JsonObject
        {
            [PendingAddinUpdate.PptkKey] = new JsonObject { ["InstalledVersion"] = installed },
            [PendingAddinUpdate.AddinKey] = new JsonObject { ["Path"] = "PPTK.ppam", ["AutoLoad"] = 1 },
            [@"Software\VB and VBA Program Settings\Instrumenta\General"] = new JsonObject { ["OperatingMode"] = "pro" },
        }.ToJsonString(), Utf8);
        SetProcesses();
        File.WriteAllText(Paths.StatePath, new JsonObject
        {
            ["schemaVersion"] = 1, ["installed"] = installed, ["previous"] = null, ["rolledBackFrom"] = null,
            ["pending"] = new JsonObject { ["version"] = pending, ["file"] = $@"pending\PPTK-{pending}.ppam", ["sha256"] = Sha(PendingFile), ["keysChanged"] = keysChanged, ["notified"] = false, ["stagedUtc"] = "2026-11-14T00:00:00Z" },
            ["legacyBackup"] = null, ["lastError"] = null, ["updatedUtc"] = "2026-11-14T00:00:00Z",
            ["architecture"] = universal ? "win-x64" : null,
        }.ToJsonString(), Utf8);
    }

    public bool Universal { get; }

    void WriteVersion(string v, string ppam, string helper)
    {
        var dir = Paths.VersionDir(v);
        Directory.CreateDirectory(Path.Combine(dir, "PPTKKeys"));
        var content = new Dictionary<string, string>
        {
            ["PPTK.ppam"] = ppam,
            [@"PPTKKeys\PPTKKeys.dll"] = helper,
            [@"PPTKKeys\PPTKKeys.deps.json"] = "{}",
            [@"PPTKKeys\PPTKKeys.runtimeconfig.json"] = "{\"runtimeOptions\":{\"framework\":{\"name\":\"Microsoft.NETCore.App\",\"version\":\"10.0.0\"}}}",
        };
        var files = new JsonArray();
        foreach (var (rel, text) in content)
        {
            var p = Path.Combine(dir, rel);
            File.WriteAllText(p, text, Utf8);
            files.Add(new JsonObject { ["path"] = "payload/" + rel.Replace('\\', '/'), ["sha256"] = Sha(p) });
        }
        if (Universal && v == Pending)
        {
            foreach (var rid in new[] { "win-x64", "win-arm64" })
                foreach (var rel in new[] { "dotnet.exe", @"host\fxr\10.0.15\hostfxr.dll", @"shared\Microsoft.NETCore.App\10.0.15\coreclr.dll" })
                {
                    string text = $"{rid} {rel}";
                    if (rid == "win-x64") { var p = Path.Combine(dir, "runtime", rel); Directory.CreateDirectory(Path.GetDirectoryName(p)!); File.WriteAllText(p, text, Utf8); }
                    using var sha = System.Security.Cryptography.SHA256.Create();
                    files.Add(new JsonObject { ["path"] = $"payload/runtime/{rid}/" + rel.Replace('\\', '/'), ["sha256"] = Convert.ToHexString(sha.ComputeHash(Utf8.GetBytes(text))).ToLowerInvariant() });
                }
        }
        File.WriteAllText(Path.Combine(dir, "version.json"), new JsonObject { ["schemaVersion"] = 1, ["packageName"] = "PPTK", ["packageVersion"] = v, ["files"] = files }.ToJsonString(), Utf8);
    }

    public string PendingFile => Path.Combine(Paths.PptkDir, "pending", $"PPTK-{Pending}.ppam");
    public string VersionAddin(string v) => Path.Combine(Paths.VersionDir(v), "PPTK.ppam");
    public void SetProcesses(params string[] names) => File.WriteAllText(Path.Combine(Root, "processes.json"), new JsonArray(names.Select(n => (JsonNode?)n).ToArray()).ToJsonString(), Utf8);
    public JsonObject State() => JsonNode.Parse(File.ReadAllText(Paths.StatePath))!.AsObject();
    public string? RegistryVersion() => (string?)JsonNode.Parse(Env.ExportRegistryKey(PendingAddinUpdate.PptkKey)!)?["InstalledVersion"];
    public PendingAddinUpdate Engine(PendingUpdateTestHooks? hooks = null, int retries = 2, List<TimeSpan>? waits = null) =>
        new(Paths, Env, retries, t => waits?.Add(t), hooks);

    public static string Sha(string path)
    {
        using var s = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Convert.ToHexString(SHA256.HashData(s)).ToLowerInvariant();
    }

    /// <summary>Every file under the root with its hash (the lock file excluded: opening it is the point).</summary>
    public string Snapshot() => string.Join("\n", Directory.GetFiles(Root, "*", SearchOption.AllDirectories)
        .Where(f => !f.EndsWith("state.lock")).OrderBy(f => f, StringComparer.Ordinal).Select(f => f[Root.Length..] + "|" + Sha(f)));

    public void Dispose() { try { Directory.Delete(Path.GetDirectoryName(Root)!, true); } catch (IOException) { } }
}

public class PendingAddinUpdateTests
{
    static void AssertInstalledOld(PendingFixture f)
    {
        Assert.Equal(PendingFixture.Sha(f.VersionAddin(f.Installed)), PendingFixture.Sha(f.Paths.AddinFile));
        Assert.Equal(f.Paths.HelperDll(f.Installed), f.Env.GetHelperTaskDll());
        Assert.Equal(f.Installed, (string?)f.State()["installed"]);
        Assert.Equal(f.Installed, f.RegistryVersion());
        Assert.False(File.Exists(f.Paths.JournalPath));
        Assert.False(Directory.Exists(f.Paths.TxnDir));
    }

    static void AssertInstalledNew(PendingFixture f, bool keysChanged)
    {
        var s = f.State();
        Assert.Equal(PendingFixture.Sha(f.VersionAddin(f.Pending)), PendingFixture.Sha(f.Paths.AddinFile));
        Assert.Equal(f.Paths.HelperDll(f.Pending), f.Env.GetHelperTaskDll());
        Assert.Equal(f.Pending, (string?)s["installed"]);
        Assert.Equal(f.Installed, (string?)s["previous"]);
        Assert.Null(s["pending"]);
        Assert.Equal(keysChanged, s["helperActivation"] is not null);
        Assert.Equal(f.Pending, f.RegistryVersion());
        Assert.False(File.Exists(f.PendingFile));
        Assert.True(Directory.Exists(f.Paths.VersionDir(f.Installed)), "the previous version is kept for rollback");
        Assert.False(File.Exists(f.Paths.JournalPath));
        Assert.False(Directory.Exists(f.Paths.TxnDir));
    }

    // ---- universal packages: the new version brings its own private runtime

    [Fact]
    public void Universal_update_switches_the_task_to_the_new_private_runtime_and_requests_a_handover()
    {
        using var f = new PendingFixture(keysChanged: true, universal: true);
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Applied, r.Outcome);
        Assert.True(r.HelperRestartRequired);
        AssertInstalledNew(f, keysChanged: true);
        Assert.Equal(f.Paths.RuntimeHost(f.Pending), f.Env.GetHelperTaskHost());
    }

    [Fact]
    public void Universal_update_with_identical_helper_files_still_needs_a_restart_for_the_new_runtime()
    {
        using var f = new PendingFixture(keysChanged: false, universal: true);   // helper DLL unchanged, host changes
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Rejected, r.Outcome);
        Assert.Contains("runtime differ", r.Detail);
        AssertInstalledOld(f);
    }

    [Theory]
    [InlineData("tamper")]
    [InlineData("missing")]
    [InlineData("architecture")]
    public void Universal_update_with_a_damaged_runtime_or_unknown_architecture_is_rejected(string damage)
    {
        using var f = new PendingFixture(keysChanged: true, universal: true);
        string coreclr = Path.Combine(f.Paths.VersionDir(f.Pending), "runtime", @"shared\Microsoft.NETCore.App\10.0.15\coreclr.dll");
        switch (damage)
        {
            case "tamper": File.AppendAllText(coreclr, "x"); break;
            case "missing": File.Delete(coreclr); break;
            case "architecture": { var s = f.State(); s["architecture"] = "win-x86"; File.WriteAllText(f.Paths.StatePath, s.ToJsonString()); break; }
        }
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Rejected, r.Outcome);
        AssertInstalledOld(f);
        Assert.NotEqual(f.Paths.RuntimeHost(f.Pending), f.Env.GetHelperTaskHost());
    }

    [Fact]
    public void Failed_universal_update_restores_the_previous_task_host()
    {
        using var f = new PendingFixture(keysChanged: true, universal: true);
        string hostBefore = f.Env.GetHelperTaskHost()!;
        var r = f.Engine(new PendingUpdateTestHooks { AfterStep = (s, _) => { if (s == "state") throw new IOException("simulated failure after state"); } }).TryApply();
        Assert.Equal(PendingUpdateOutcome.RolledBack, r.Outcome);
        AssertInstalledOld(f);
        Assert.Equal(hostBefore, f.Env.GetHelperTaskHost());
    }

    [Fact]
    public void Valid_pending_update_with_unchanged_helper_is_applied_without_restart()
    {
        using var f = new PendingFixture(keysChanged: false);
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Applied, r.Outcome);
        Assert.False(r.HelperRestartRequired);
        AssertInstalledNew(f, keysChanged: false);
    }

    [Fact]
    public void KeysChanged_update_repoints_the_task_and_requests_a_handover_without_touching_the_running_helper()
    {
        using var f = new PendingFixture(keysChanged: true);
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Applied, r.Outcome);
        Assert.True(r.HelperRestartRequired);
        Assert.Equal(f.Paths.HelperDll(f.Pending), r.HelperDll);
        AssertInstalledNew(f, keysChanged: true);
        Assert.Equal(f.Pending, (string?)f.State()["helperActivation"]!["version"]);
        Assert.True(File.Exists(f.Paths.HelperDll(f.Installed)), "the running helper's files stay in place");
    }

    [Fact]
    public void Incorrect_sha256_is_rejected_and_nothing_is_installed()
    {
        using var f = new PendingFixture();
        File.AppendAllText(f.PendingFile, "tampered");
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Rejected, r.Outcome);
        Assert.Contains("SHA-256", r.Detail);
        AssertInstalledOld(f);
        Assert.NotNull(f.State()["pending"]);
        Assert.Contains("SHA-256", (string?)f.State()["lastError"]);
    }

    [Fact]
    public void Missing_pending_ppam_is_rejected()
    {
        using var f = new PendingFixture();
        File.Delete(f.PendingFile);
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Rejected, r.Outcome);
        Assert.Contains("missing", r.Detail);
        AssertInstalledOld(f);
    }

    [Fact]
    public void PowerPoint_still_running_defers_and_changes_nothing()
    {
        using var f = new PendingFixture();
        f.SetProcesses("POWERPNT");
        var before = f.Snapshot();
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Deferred, r.Outcome);
        Assert.Equal(before, f.Snapshot());
    }

    [Fact]
    public void Locked_destination_is_retried_then_deferred_without_changes()
    {
        using var f = new PendingFixture();
        var before = f.Snapshot();
        var waits = new List<TimeSpan>();
        PendingUpdateResult r;
        using (new FileStream(f.Paths.AddinFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            r = f.Engine(retries: 3, waits: waits).TryApply();
        Assert.Equal(PendingUpdateOutcome.Deferred, r.Outcome);
        Assert.Contains("in use", r.Detail);
        Assert.Equal(3, waits.Count);
        Assert.Equal(before, f.Snapshot());
        Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);   // released: applies now
    }

    [Fact]
    public void PowerPoint_starting_while_waiting_for_the_lock_defers()
    {
        using var f = new PendingFixture();
        var before = f.Snapshot();
        PendingUpdateResult r;
        using (new FileStream(f.Paths.AddinFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            r = new PendingAddinUpdate(f.Paths, f.Env, 3, _ => f.SetProcesses("POWERPNT")).TryApply();
        Assert.Equal(PendingUpdateOutcome.Deferred, r.Outcome);
        Assert.Contains("PowerPoint is running", r.Detail);
        f.SetProcesses();
        Assert.Equal(before, f.Snapshot());
    }

    [Fact]
    public void Concurrent_installer_holding_state_lock_defers()
    {
        using var f = new PendingFixture();
        var before = f.Snapshot();
        PendingUpdateResult r;
        // Exactly how install.ps1 / uninstall.ps1 hold it.
        using (new FileStream(f.Paths.LockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Deferred, r.Outcome);
        Assert.Contains("state.lock", r.Detail);
        Assert.Equal(before, f.Snapshot());
    }

    [Fact]
    public void Interrupted_installer_operation_is_left_to_the_installer()
    {
        using var f = new PendingFixture();
        File.WriteAllText(f.Paths.JournalPath, "{\"operation\":\"install\",\"target\":\"1.0.1\"}");
        var before = f.Snapshot();
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Deferred, r.Outcome);
        Assert.Contains("install.ps1", r.Detail);
        Assert.Equal(before, f.Snapshot());
    }

    [Theory]
    [InlineData("{ not json")]
    [InlineData("[1,2]")]
    public void Corrupt_state_json_is_rejected_and_left_untouched(string text)
    {
        using var f = new PendingFixture();
        File.WriteAllText(f.Paths.StatePath, text);
        var before = f.Snapshot();
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Rejected, r.Outcome);
        Assert.Equal(before, f.Snapshot());
    }

    [Theory]
    [InlineData("file", "..\\..\\AppData\\Microsoft\\AddIns\\PPTK.ppam", "pending file path")]
    [InlineData("version", "0.9.0", "not newer")]
    [InlineData("version", "1.0.0", "not newer")]
    [InlineData("version", "latest", "valid version")]
    [InlineData("sha256", "ABC", "valid sha256")]
    public void Invalid_pending_entries_are_rejected(string field, string value, string expected)
    {
        using var f = new PendingFixture();
        var s = f.State(); s["pending"]![field] = value;
        File.WriteAllText(f.Paths.StatePath, s.ToJsonString());
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Rejected, r.Outcome);
        Assert.Contains(expected, r.Detail);
        AssertInstalledOld(f);
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void KeysChanged_flag_that_does_not_match_the_files_is_rejected(bool actual, bool claimed)
    {
        using var f = new PendingFixture(keysChanged: actual);
        var s = f.State(); s["pending"]!["keysChanged"] = claimed;
        File.WriteAllText(f.Paths.StatePath, s.ToJsonString());
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Rejected, r.Outcome);
        Assert.Contains("keysChanged", r.Detail);
        AssertInstalledOld(f);
    }

    [Fact]
    public void Missing_keysChanged_flag_is_rejected()
    {
        using var f = new PendingFixture();
        var s = f.State(); s["pending"]!.AsObject().Remove("keysChanged");
        File.WriteAllText(f.Paths.StatePath, s.ToJsonString());
        Assert.Equal(PendingUpdateOutcome.Rejected, f.Engine().TryApply().Outcome);
        AssertInstalledOld(f);
    }

    [Theory]
    [InlineData("helper")]
    [InlineData("addin")]
    [InlineData("registry")]
    [InlineData("newfiles")]
    public void Unexpected_installation_is_rejected(string what)
    {
        using var f = new PendingFixture();
        switch (what)
        {
            case "helper": f.Env.SetHelperTaskDll(@"C:\Elsewhere\PPTKKeys.dll"); break;
            case "addin": File.WriteAllText(f.Paths.AddinFile, "someone else's add-in"); break;
            case "registry": f.Env.ImportRegistryKey(PendingAddinUpdate.PptkKey, "{\"InstalledVersion\":\"0.9.9\"}"); break;
            case "newfiles": File.Delete(Path.Combine(f.Paths.VersionDir(f.Pending), "PPTKKeys", "PPTKKeys.deps.json")); break;
        }
        var before = f.Snapshot().Split('\n').Where(l => !l.StartsWith(@"\LocalAppData\PPTK\state.json")).ToArray();
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Rejected, r.Outcome);
        Assert.Equal(before, f.Snapshot().Split('\n').Where(l => !l.StartsWith(@"\LocalAppData\PPTK\state.json")).ToArray());
        Assert.Equal(f.Installed, (string?)f.State()["installed"]);
    }

    [Theory]
    [InlineData("journal")]
    [InlineData("replaced")]
    [InlineData("task")]
    [InlineData("state")]
    [InlineData("registry")]
    public void Interrupted_update_is_undone_by_the_next_run_which_then_applies_it(string step)
    {
        using var f = new PendingFixture(keysChanged: true);
        var r = f.Engine(new PendingUpdateTestHooks { AbandonAfter = s => s == step }).TryApply();
        Assert.Equal(PendingUpdateOutcome.Interrupted, r.Outcome);
        Assert.True(File.Exists(f.Paths.JournalPath));
        var again = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.Applied, again.Outcome);
        AssertInstalledNew(f, keysChanged: true);
    }

    [Fact]
    public void Failed_replacement_restores_the_installed_version()
    {
        using var f = new PendingFixture();
        var r = f.Engine(new PendingUpdateTestHooks { AfterStep = (s, file) => { if (s == "copied") File.Delete(file!); } }).TryApply();
        Assert.Equal(PendingUpdateOutcome.RolledBack, r.Outcome);
        AssertInstalledOld(f);
        Assert.NotNull(f.State()["pending"]);
        Assert.NotNull((string?)f.State()["lastError"]);
    }

    [Fact]
    public void Failed_verification_after_replacing_restores_the_previous_add_in()
    {
        using var f = new PendingFixture();
        var r = f.Engine(new PendingUpdateTestHooks { AfterStep = (s, file) => { if (s == "replaced") File.AppendAllText(file!, "corrupted"); } }).TryApply();
        Assert.Equal(PendingUpdateOutcome.RolledBack, r.Outcome);
        Assert.Contains("does not match", r.Detail);
        AssertInstalledOld(f);
    }

    [Theory]
    [InlineData("task")]
    [InlineData("state")]
    [InlineData("registry")]
    public void Failure_after_later_steps_rolls_everything_back(string step)
    {
        using var f = new PendingFixture(keysChanged: true);
        var tasks = f.Env.ExportTasks();
        var r = f.Engine(new PendingUpdateTestHooks { AfterStep = (s, _) => { if (s == step) throw new IOException("simulated failure"); } }).TryApply();
        Assert.Equal(PendingUpdateOutcome.RolledBack, r.Outcome);
        AssertInstalledOld(f);
        Assert.Equal(tasks, f.Env.ExportTasks());
        Assert.Null(f.State()["helperActivation"]);
        Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);   // the update stays pending and applies later
        AssertInstalledNew(f, keysChanged: true);
    }

    [Fact]
    public void Restore_failure_is_reported_and_the_journal_kept()
    {
        using var f = new PendingFixture();
        var r = f.Engine(new PendingUpdateTestHooks
        {
            AfterStep = (s, _) =>
            {
                if (s != "replaced") return;
                File.AppendAllText(Path.Combine(f.Paths.TxnDir, "PPTK.ppam.before"), "x");
                throw new IOException("simulated failure");
            },
        }).TryApply();
        Assert.Equal(PendingUpdateOutcome.RestoreFailed, r.Outcome);
        Assert.Contains("RESTORE FAILED", r.Detail);
        Assert.True(File.Exists(f.Paths.JournalPath));
    }

    [Fact]
    public void Clean_up_failure_after_the_commit_never_rolls_back_the_applied_update()
    {
        using var f = new PendingFixture(keysChanged: true);
        FileStream? held = null;
        // Something else keeps the journal's add-in copy open (no delete sharing): txn\ cannot be deleted.
        var r = f.Engine(new PendingUpdateTestHooks { AfterStep = (s, _) => { if (s == "registry") held = new FileStream(Path.Combine(f.Paths.TxnDir, "PPTK.ppam.before"), FileMode.Open, FileAccess.Read, FileShare.Read); } }).TryApply();
        try
        {
            Assert.True(r.Outcome == PendingUpdateOutcome.Applied, $"{r.Outcome}: {r.Detail}");
            Assert.False(File.Exists(f.Paths.JournalPath));
            Assert.True(Directory.Exists(f.Paths.TxnDir), "the clean-up really failed");
            Assert.Equal(f.Pending, (string?)f.State()["installed"]);
            Assert.Equal(PendingFixture.Sha(f.VersionAddin(f.Pending)), PendingFixture.Sha(f.Paths.AddinFile));
        }
        finally { held?.Dispose(); }
        // The next run removes the leftover txn\ and changes nothing else.
        Assert.Equal(PendingUpdateOutcome.NothingPending, f.Engine().TryApply().Outcome);
        AssertInstalledNew(f, keysChanged: true);
    }

    // ---- bounded retry of replace operations
    static IOException CannotRemoveReplaced() => new("simulated: Unable to remove the file to be replaced", unchecked((int)0x80070497));
    static IOException SharingViolation() => new("simulated: sharing violation", unchecked((int)0x80070020));

    static PendingUpdateTestHooks FailFirst(string operation, int times, Func<Exception> make, Action<string, string>? before = null, List<string>? calls = null)
    {
        int left = times;
        return new PendingUpdateTestHooks
        {
            FailOperation = (op, source, target) =>
            {
                calls?.Add(op);
                if (op != operation || left <= 0) return null;
                left--;
                before?.Invoke(source, target);
                return make();
            },
        };
    }

    [Theory]
    [InlineData("replace state.json")]
    [InlineData("replace PPTK.ppam")]
    [InlineData("replace transaction.json")]
    public void Transient_replace_failures_are_retried_and_the_update_applies(string operation)
    {
        using var f = new PendingFixture(keysChanged: true);   // (transaction.json is new: its write is a move, retried the same way)
        var calls = new List<string>();
        var engine = f.Engine(FailFirst(operation, 3, CannotRemoveReplaced, calls: calls));
        var r = engine.TryApply();
        Assert.True(r.Outcome == PendingUpdateOutcome.Applied, $"{r.Outcome}: {r.Detail}");
        AssertInstalledNew(f, keysChanged: true);
        Assert.Equal(4, calls.Count(c => c == operation));
        Assert.Contains(engine.Log, l => l.Contains($"{operation}: attempt 3 failed (0x80070497"));
    }

    [Fact]
    public void Retry_exhaustion_on_the_add_in_rolls_back_and_keeps_the_update_pending()
    {
        using var f = new PendingFixture(keysChanged: true);
        var calls = new List<string>();
        var engine = f.Engine(FailFirst("replace PPTK.ppam", 5, SharingViolation, calls: calls));
        var r = engine.TryApply();
        Assert.True(r.Outcome == PendingUpdateOutcome.RolledBack, $"{r.Outcome}: {r.Detail}");
        Assert.Contains("sharing violation", r.Detail);
        Assert.Equal(5, calls.Count(c => c == "replace PPTK.ppam"));
        Assert.Contains(engine.Log, l => l.Contains("attempt 5 of 5 failed; giving up"));
        AssertInstalledOld(f);
        Assert.NotNull(f.State()["pending"]);
        Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);   // a later pass applies it
    }

    [Fact]
    public void PowerPoint_starting_between_attempts_stops_the_add_in_replacement()
    {
        using var f = new PendingFixture();
        var calls = new List<string>();
        var r = f.Engine(FailFirst("replace PPTK.ppam", 1, SharingViolation, before: (_, _) => f.SetProcesses("POWERPNT"), calls: calls)).TryApply();
        Assert.True(r.Outcome == PendingUpdateOutcome.RolledBack, $"{r.Outcome}: {r.Detail}");
        Assert.Contains("PowerPoint started", r.Detail);
        Assert.Equal(1, calls.Count(c => c == "replace PPTK.ppam"));   // no second attempt under PowerPoint
        f.SetProcesses();
        AssertInstalledOld(f);
    }

    [Fact]
    public void Access_denied_on_a_read_only_target_is_permanent_and_not_retried()
    {
        using var f = new PendingFixture();
        File.SetAttributes(f.Paths.AddinFile, FileAttributes.ReadOnly);
        try
        {
            var calls = new List<string>();
            var r = f.Engine(FailFirst("replace PPTK.ppam", 5, () => new UnauthorizedAccessException("simulated: access denied"), calls: calls)).TryApply();
            Assert.True(r.Outcome == PendingUpdateOutcome.RolledBack, $"{r.Outcome}: {r.Detail}");
            Assert.Equal(1, calls.Count(c => c == "replace PPTK.ppam"));
        }
        finally { File.SetAttributes(f.Paths.AddinFile, FileAttributes.Normal); }
        AssertInstalledOld(f);
    }

    [Fact]
    public void Access_denied_on_a_writable_target_is_retried_at_most_three_times()
    {
        using var f = new PendingFixture();
        var engine = f.Engine(FailFirst("replace state.json", 3, () => new UnauthorizedAccessException("simulated: access denied")));
        var r = engine.TryApply();
        Assert.True(r.Outcome == PendingUpdateOutcome.RolledBack, $"{r.Outcome}: {r.Detail}");
        Assert.Contains(engine.Log, l => l.Contains("replace state.json: attempt 3 of 3 failed; giving up"));   // the restore's own writes then succeed
        AssertInstalledOld(f);
    }

    [Fact]
    public void An_attempt_that_completed_but_reported_failure_is_not_repeated()
    {
        using var f = new PendingFixture();
        var calls = new List<string>();
        // The replace happens, then an error is reported anyway (partial success).
        var engine = f.Engine(FailFirst("replace state.json", 1, CannotRemoveReplaced, before: (source, target) => File.Replace(source, target, null), calls: calls));
        var r = engine.TryApply();
        Assert.True(r.Outcome == PendingUpdateOutcome.Applied, $"{r.Outcome}: {r.Detail}");
        Assert.Contains(engine.Log, l => l.Contains("replace state.json: the earlier attempt had completed"));
        Assert.Equal(1, calls.Count(c => c == "replace state.json"));
        AssertInstalledNew(f, keysChanged: false);
    }

    [Fact]
    public void Retry_classification()
    {
        var dir = Path.Combine(Path.GetTempPath(), "PPTK-retry-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(dir);
        var file = Path.Combine(dir, "x.json"); File.WriteAllText(file, "{}");
        try
        {
            Assert.Equal(5, PendingAddinUpdate.RetryAttempts(SharingViolation(), file));
            Assert.Equal(5, PendingAddinUpdate.RetryAttempts(new IOException("lock", unchecked((int)0x80070021)), file));
            Assert.Equal(5, PendingAddinUpdate.RetryAttempts(CannotRemoveReplaced(), file));
            Assert.Equal(5, PendingAddinUpdate.RetryAttempts(new IOException("move", unchecked((int)0x80070498)), file));
            Assert.Equal(3, PendingAddinUpdate.RetryAttempts(new UnauthorizedAccessException("denied"), file));
            Assert.Equal(3, PendingAddinUpdate.RetryAttempts(new IOException("denied", unchecked((int)0x80070005)), file));
            Assert.Equal(1, PendingAddinUpdate.RetryAttempts(new IOException("disk full", unchecked((int)0x80070070)), file));
            Assert.Equal(1, PendingAddinUpdate.RetryAttempts(new FileNotFoundException("gone"), file));
            File.SetAttributes(file, FileAttributes.ReadOnly);
            Assert.Equal(1, PendingAddinUpdate.RetryAttempts(new UnauthorizedAccessException("denied"), file));
        }
        finally { File.SetAttributes(file, FileAttributes.Normal); Directory.Delete(dir, true); }
    }

    [Fact]
    public void Repeated_execution_is_a_no_op()
    {
        using var f = new PendingFixture();
        Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);
        var before = f.Snapshot();
        var r = f.Engine().TryApply();
        Assert.Equal(PendingUpdateOutcome.NothingPending, r.Outcome);
        Assert.Equal(before, f.Snapshot());
    }

    [Fact]
    public void No_installation_is_nothing_pending()
    {
        var dir = Path.Combine(Path.GetTempPath(), "PPTK-pending-empty-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            var r = new PendingAddinUpdate(new PendingUpdatePaths(Path.Combine(dir, "AppData"), Path.Combine(dir, "LocalAppData")), new SimulatedEnvironment(dir)).TryApply();
            Assert.Equal(PendingUpdateOutcome.NothingPending, r.Outcome);
            Assert.False(Directory.Exists(dir));
        }
        finally { if (Directory.Exists(dir)) Directory.Delete(dir, true); }
    }

    [Fact]
    public void Unicode_paths_work()
    {
        using var f = new PendingFixture(keysChanged: true, name: "اختبار ✓ PPTK");
        Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);
        AssertInstalledNew(f, keysChanged: true);
    }

    [Fact]
    public void No_changes_outside_the_test_directory()
    {
        static string Live()
        {
            var addin = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\AddIns\PPTK.ppam");
            var local = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PPTK");
            var parts = new List<string> { File.Exists(addin) ? PendingFixture.Sha(addin) : "-" };
            if (Directory.Exists(local)) parts.AddRange(Directory.GetFiles(local, "*", SearchOption.AllDirectories).OrderBy(x => x).Select(x => x + "|" + new FileInfo(x).Length + "|" + File.GetLastWriteTimeUtc(x).Ticks));
            return string.Join("\n", parts);
        }
        var before = Live();
        using (var f = new PendingFixture(keysChanged: true))
        {
            Assert.Equal(PendingUpdateOutcome.Applied, f.Engine().TryApply().Outcome);
            Assert.StartsWith(Path.GetTempPath(), f.Paths.AddinFile);
        }
        Assert.Equal(before, Live());
    }
}

/// <summary>The engine against the real src\Deploy scripts (install.ps1 stages, detect.ps1 verifies).</summary>
public class PendingAddinUpdateScriptTests
{
    static readonly UTF8Encoding Utf8 = new(false);

    internal static string Deploy()
    {
        for (var d = new DirectoryInfo(AppContext.BaseDirectory); d is not null; d = d.Parent)
            if (File.Exists(Path.Combine(d.FullName, @"src\Deploy\install.ps1"))) return Path.Combine(d.FullName, @"src\Deploy");
        throw new InvalidOperationException("src\\Deploy not found");
    }

    internal static (int Code, string Out, string Err) Ps(string script, params string[] args)
    {
        var psi = new ProcessStartInfo("powershell.exe") { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
        foreach (var a in new[] { "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", Path.Combine(Deploy(), script) }.Concat(args)) psi.ArgumentList.Add(a);
        using var p = Process.Start(psi)!;
        var err = p.StandardError.ReadToEndAsync();   // both streams at once: a full stderr pipe cannot block the child
        string o = p.StandardOutput.ReadToEnd(), e = err.GetAwaiter().GetResult();
        p.WaitForExit();
        return (p.ExitCode, o.Trim(), e.Trim());
    }

    /// <summary>The scripts' own logs under a test root, for failure messages.</summary>
    internal static string Logs(string root)
    {
        var dir = Path.Combine(root, @"LocalAppData\PPTK\logs");
        if (!Directory.Exists(dir)) return "(no logs)";
        return string.Join("\n", Directory.GetFiles(dir).OrderBy(f => f, StringComparer.Ordinal).Select(f => "--- " + Path.GetFileName(f) + "\n" + File.ReadAllText(f)));
    }

    internal static string Package(string baseDir, string version, string addin, string helper)
    {
        var dir = Path.Combine(baseDir, "pkg-" + version);
        Directory.CreateDirectory(Path.Combine(dir, @"payload\PPTKKeys"));
        var files = new JsonArray();
        foreach (var (rel, text) in new[] { ("payload/PPTK.ppam", "add-in " + addin), ("payload/PPTKKeys/PPTKKeys.dll", "helper " + helper), ("payload/PPTKKeys/PPTKKeys.deps.json", "{}"), ("payload/PPTKKeys/PPTKKeys.runtimeconfig.json", "{\"runtimeOptions\":{\"framework\":{\"name\":\"Microsoft.NETCore.App\",\"version\":\"10.0.0\"}}}") })
        {
            var p = Path.Combine(dir, rel.Replace('/', '\\'));
            File.WriteAllText(p, text, Utf8);
            files.Add(new JsonObject { ["path"] = rel, ["sha256"] = PendingFixture.Sha(p) });
        }
        File.WriteAllText(Path.Combine(dir, "version.json"), new JsonObject { ["schemaVersion"] = 1, ["packageName"] = "PPTK", ["packageVersion"] = version, ["files"] = files }.ToJsonString(), Utf8);
        return dir;
    }

    internal static string Root(string baseDir)
    {
        var r = Path.Combine(baseDir, "root");
        Directory.CreateDirectory(Path.Combine(r, @"LocalAppData\Microsoft\dotnet\shared\Microsoft.NETCore.App\10.0.14"));
        File.WriteAllText(Path.Combine(r, @"LocalAppData\Microsoft\dotnet\dotnet.exe"), "placeholder; never executed", Utf8);
        File.WriteAllText(Path.Combine(r, "PPTK-TEST-ROOT"), "isolated", Utf8);
        File.WriteAllText(Path.Combine(r, "identity.json"), "{\"name\":\"TESTDOM\\\\tester\",\"sid\":\"S-1-5-21-1\"}", Utf8);
        File.WriteAllText(Path.Combine(r, "registry.json"), "{}", Utf8);
        File.WriteAllText(Path.Combine(r, "processes.json"), "[]", Utf8);
        return r;
    }

    [Fact]
    public void Update_staged_by_install_ps1_is_applied_and_confirmed_by_detect_ps1()
    {
        var b = Path.Combine(Path.GetTempPath(), "PPTK-pending-scripts-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            var p100 = Package(b, "1.0.0", "A", "A"); var p101 = Package(b, "1.0.1", "B", "B"); var r = Root(b);
            Assert.Equal(0, Ps("install.ps1", "-Root", r, "-PackagePath", p100).Code);
            File.WriteAllText(Path.Combine(r, "processes.json"), "[\"POWERPNT\"]", Utf8);
            Assert.Equal(1618, Ps("install.ps1", "-Root", r, "-PackagePath", p101).Code);
            Assert.Equal(0, Ps("detect.ps1", "-Root", r, "-ManifestPath", Path.Combine(p100, "version.json")).Code);
            File.WriteAllText(Path.Combine(r, "processes.json"), "[]", Utf8);

            var res = new PendingAddinUpdate(new PendingUpdatePaths(Path.Combine(r, "AppData"), Path.Combine(r, "LocalAppData")), new SimulatedEnvironment(r)).TryApply();
            Assert.Equal(PendingUpdateOutcome.Applied, res.Outcome);
            Assert.True(res.HelperRestartRequired);

            var d = Ps("detect.ps1", "-Root", r, "-ManifestPath", Path.Combine(p101, "version.json"));
            Assert.True(d.Code == 0, d.Err);
            Assert.NotEqual(0, Ps("detect.ps1", "-Root", r, "-ManifestPath", Path.Combine(p100, "version.json")).Code);
            var again = Ps("install.ps1", "-Root", r, "-PackagePath", p101);
            Assert.Equal(0, again.Code);
            Assert.Contains("no change", again.Out);
            Assert.Equal(0, Ps("uninstall.ps1", "-Root", r).Code);
        }
        finally { if (Directory.Exists(b)) Directory.Delete(b, true); }
    }

    [Fact]
    public void Interrupted_engine_run_is_undone_by_install_ps1()
    {
        var b = Path.Combine(Path.GetTempPath(), "PPTK-pending-scripts-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            var p100 = Package(b, "1.0.0", "A", "A"); var p101 = Package(b, "1.0.1", "B", "B"); var r = Root(b);
            Assert.Equal(0, Ps("install.ps1", "-Root", r, "-PackagePath", p100).Code);
            File.WriteAllText(Path.Combine(r, "processes.json"), "[\"POWERPNT\"]", Utf8);
            Assert.Equal(1618, Ps("install.ps1", "-Root", r, "-PackagePath", p101).Code);
            File.WriteAllText(Path.Combine(r, "processes.json"), "[]", Utf8);

            var res = new PendingAddinUpdate(new PendingUpdatePaths(Path.Combine(r, "AppData"), Path.Combine(r, "LocalAppData")), new SimulatedEnvironment(r),
                hooks: new PendingUpdateTestHooks { AbandonAfter = s => s == "task" }).TryApply();
            Assert.True(res.Outcome == PendingUpdateOutcome.Interrupted, $"{res.Outcome}: {res.Detail}");

            // The installer's retry finds the engine's journal, undoes it (same format), then installs 1.0.1 itself.
            var i = Ps("install.ps1", "-Root", r, "-PackagePath", p101);
            Assert.True(i.Code == 0, i.Out + " " + i.Err);
            Assert.Equal(0, Ps("detect.ps1", "-Root", r, "-ManifestPath", Path.Combine(p101, "version.json")).Code);
            Assert.False(File.Exists(Path.Combine(r, @"LocalAppData\PPTK\transaction.json")));
        }
        finally { if (Directory.Exists(b)) Directory.Delete(b, true); }
    }
}
