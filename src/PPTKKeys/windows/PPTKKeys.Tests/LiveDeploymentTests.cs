using System.Diagnostics;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text.Json.Nodes;
using Microsoft.Win32;
using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

/// <summary>Read-only snapshot of the production PPTK state; every live test compares it before and after.</summary>
[SupportedOSPlatform("windows")]
sealed class ProductionSnapshot
{
    public string Text { get; }

    ProductionSnapshot(string text) { Text = text; }

    public static ProductionSnapshot Take()
    {
        var lines = new List<string>();
        dynamic service = Activator.CreateInstance(Type.GetTypeFromProgID("Schedule.Service", true)!)!;
        service.Connect();
        try { lines.Add("task=" + (string)service.GetFolder("\\").GetTask(LiveDeploymentConfig.ProductionTaskName).Xml); }
        catch (Exception ex) when (LiveDeploymentEnvironment.IsNotFound(ex)) { lines.Add("task=(none)"); }
        foreach (var key in new[] { PendingAddinUpdate.PptkKey, PendingAddinUpdate.AddinKey })
        {
            using var k = Registry.CurrentUser.OpenSubKey(key);
            lines.Add($"{key}=" + (k is null ? "(none)" : string.Join(";", k.GetValueNames().OrderBy(n => n).Select(n => $"{n}:{k.GetValueKind(n)}:{k.GetValue(n)}"))));
        }
        lines.Add("ppam=" + Hash(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\AddIns\PPTK.ppam")));
        lines.Add("dll=" + Hash(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"PPTK\PPTKKeys\PPTKKeys.dll")));
        // The running helper holds this mutex; opening it (never acquiring) only checks that it still runs.
        bool helper = Mutex.TryOpenExisting(@"Local\PPTKKeys.Windows", out var m);
        m?.Dispose();
        lines.Add("helper running=" + helper);
        return new ProductionSnapshot(string.Join("\n", lines));
    }

    static string Hash(string path)
    {
        if (!File.Exists(path)) return "(none)";
        using var s = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Convert.ToHexString(SHA256.HashData(s));
    }
}

/// <summary>One disabled, trigger-less test task "\PPTK-Test-&lt;hex&gt;" and one key HKCU\Software\PPTK-Test-&lt;hex&gt;,
/// both created by this instance and removed only by it. A failed clean-up names the exact leftovers.</summary>
[SupportedOSPlatform("windows")]
sealed class TestResources : IDisposable
{
    public string Id { get; } = "PPTK-Test-" + Guid.NewGuid().ToString("N")[..12];
    public string RegistryRoot => @"Software\" + Id;
    public LiveDeploymentConfig Config { get; }
    public FakeProbe Probe { get; } = new();
    public LiveDeploymentEnvironment Env { get; }
    readonly ProductionSnapshot before;
    bool taskCreated;

    public TestResources()
    {
        // Validated before anything is created: test names only, nothing of these names exists yet.
        Config = LiveDeploymentConfig.ForTest(Id, RegistryRoot);
        before = ProductionSnapshot.Take();
        using (var k = Registry.CurrentUser.OpenSubKey(RegistryRoot)) if (k is not null) throw new InvalidOperationException($"{RegistryRoot} already exists");
        Env = new LiveDeploymentEnvironment(Config, Probe);
        if (Env.GetHelperTaskDll() is not null || Env.ExportTasks() is not null) throw new InvalidOperationException($@"task \{Id} already exists");
    }

    public static string Dotnet => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\dotnet\dotnet.exe");

    /// <summary>Creates the task: disabled, no trigger, current user, interactive token, least privilege.</summary>
    public void CreateTask(string dll)
    {
        dynamic service = Activator.CreateInstance(Type.GetTypeFromProgID("Schedule.Service", true)!)!;
        service.Connect();
        dynamic def = service.NewTask(0);
        def.RegistrationInfo.Description = "PPTK automated test task (disabled, never runs). Safe to delete.";
        def.Settings.Enabled = false;
        def.Settings.ExecutionTimeLimit = "PT0S";
        def.Settings.MultipleInstances = 2;   // IgnoreNew
        using var me = WindowsIdentity.GetCurrent();
        def.Principal.UserId = me.Name;
        def.Principal.LogonType = 3;   // interactive token
        def.Principal.RunLevel = 0;    // least privilege
        dynamic action = def.Actions.Create(0);
        action.Path = Dotnet;
        action.Arguments = $"exec \"{dll}\"";
        action.WorkingDirectory = Path.GetDirectoryName(dll);
        taskCreated = true;
        service.GetFolder("\\").RegisterTaskDefinition(Id, def, 6, null, null, 3, null);
    }

    public (bool Enabled, int Triggers, string User, int Logon, int RunLevel) TaskInfo()
    {
        dynamic service = Activator.CreateInstance(Type.GetTypeFromProgID("Schedule.Service", true)!)!;
        service.Connect();
        dynamic t = service.GetFolder("\\").GetTask(Id);
        dynamic d = t.Definition;
        return ((bool)t.Enabled, (int)d.Triggers.Count, (string)d.Principal.UserId, (int)d.Principal.LogonType, (int)d.Principal.RunLevel);
    }

    public void AssertProductionUnchanged() => Assert.Equal(before.Text, ProductionSnapshot.Take().Text);

    public void Dispose()
    {
        var leftovers = new List<string>();
        try
        {
            dynamic service = Activator.CreateInstance(Type.GetTypeFromProgID("Schedule.Service", true)!)!;
            service.Connect();
            dynamic root = service.GetFolder("\\");
            try { root.DeleteTask(Id, 0); } catch (Exception ex) when (LiveDeploymentEnvironment.IsNotFound(ex)) { }
            try { root.GetTask(Id); leftovers.Add($@"scheduled task \{Id}"); } catch (Exception ex) when (LiveDeploymentEnvironment.IsNotFound(ex)) { }
        }
        catch (Exception ex) when (taskCreated) { leftovers.Add($@"scheduled task \{Id} (clean-up error: {ex.Message})"); }
        try { Registry.CurrentUser.DeleteSubKeyTree(RegistryRoot, throwOnMissingSubKey: false); } catch (Exception ex) { leftovers.Add($@"HKCU\{RegistryRoot} ({ex.Message})"); }
        using (var k = Registry.CurrentUser.OpenSubKey(RegistryRoot)) if (k is not null && !leftovers.Any(l => l.Contains(RegistryRoot))) leftovers.Add($@"HKCU\{RegistryRoot}");
        if (leftovers.Count > 0) throw new InvalidOperationException("TEST CLEAN-UP FAILED, leftovers: " + string.Join("; ", leftovers));
    }
}

[SupportedOSPlatform("windows")]
public class LiveDeploymentTests
{
    [Fact]
    public void Test_configurations_refuse_production_names_and_production_is_a_separate_configuration()
    {
        Assert.True(DeploymentFlags.ProductionUpdatesEnabled);
        Assert.False(DeploymentFlags.ForcedRecoveryEnabled);
        var production = LiveDeploymentConfig.Production();
        Assert.True(production.IsProduction);
        Assert.Equal("PPTKKeys Startup", production.TaskName);
        Assert.Equal(@"Software\PPTK", production.MapKey(@"Software\PPTK"));
        Assert.Throws<InvalidOperationException>(() => production.CheckTask("\\", "PPTK-Test-0123abcd"));
        Assert.Throws<ArgumentException>(() => LiveDeploymentConfig.ForTest("PPTKKeys Startup", @"Software\PPTK-Test-0123abcd"));
        Assert.Throws<ArgumentException>(() => LiveDeploymentConfig.ForTest("PPTK-Test-0123abcd", @"Software\PPTK"));
        Assert.Throws<ArgumentException>(() => LiveDeploymentConfig.ForTest("PPTK-Test-0123abcd", @"Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK"));
        Assert.Throws<ArgumentException>(() => LiveDeploymentConfig.ForTest("PPTK-Test-0123abcd", @"Software\PPTK-Test-0123abcd\Sub"));
        Assert.Throws<ArgumentException>(() => LiveDeploymentConfig.ForTest(@"..\PPTKKeys Startup", @"Software\PPTK-Test-0123abcd"));
        var c = LiveDeploymentConfig.ForTest("PPTK-Test-0123abcd", @"Software\PPTK-Test-0123abcd");
        Assert.Equal(@"Software\PPTK-Test-0123abcd\Software\PPTK", c.MapKey(PendingAddinUpdate.PptkKey));
        Assert.Throws<InvalidOperationException>(() => c.CheckTask("\\", "PPTKKeys Startup"));
        Assert.Throws<InvalidOperationException>(() => c.CheckTask(@"\Other\", "PPTK-Test-0123abcd"));
        Assert.Throws<ArgumentException>(() => c.MapKey(@"Software\Something"));
    }

    [Fact]
    public void Task_round_trip_changes_only_the_action_and_restores_exactly()
    {
        using var t = new TestResources();
        string a = @"C:\PPTK-Test\versions\1.0.0\PPTKKeys\PPTKKeys.dll", b = @"C:\PPTK-Test\versions\1.0.1\PPTKKeys\PPTKKeys.dll";
        t.CreateTask(a);
        var info = t.TaskInfo();
        Assert.Equal(a, t.Env.GetHelperTaskDll());
        var snapshot = t.Env.ExportTasks()!;
        Assert.Equal(t.Id, (string?)JsonNode.Parse(snapshot)!["tasks"]![0]!["name"]);

        t.Env.SetHelperTaskDll(b);
        Assert.Equal(b, t.Env.GetHelperTaskDll());
        Assert.Equal(info, t.TaskInfo());   // still disabled, no trigger, same user, interactive, least privilege
        Assert.False(info.Enabled);
        Assert.Equal(0, info.Triggers);

        // Universal version: the program switches to that version's private runtime as well.
        string hostBefore = t.Env.GetHelperTaskHost()!, host = @"C:\PPTK-Test\versions\1.0.1\runtime\dotnet.exe";
        t.Env.SetHelperTaskDll(b, host);
        Assert.Equal(host, t.Env.GetHelperTaskHost());
        Assert.Equal(b, t.Env.GetHelperTaskDll());
        Assert.Equal(info, t.TaskInfo());

        t.Env.ImportTasks(snapshot);
        Assert.Equal(a, t.Env.GetHelperTaskDll());
        Assert.Equal(hostBefore, t.Env.GetHelperTaskHost());
        Assert.Equal(info, t.TaskInfo());
        t.AssertProductionUnchanged();
    }

    [Fact]
    public void Null_snapshot_removes_only_the_test_task()
    {
        using var t = new TestResources();
        t.CreateTask(@"C:\PPTK-Test\x\PPTKKeys.dll");
        t.Env.ImportTasks(null);
        Assert.Null(t.Env.ExportTasks());
        t.Env.ImportTasks(null);   // idempotent
        t.AssertProductionUnchanged();
    }

    [Fact]
    public void Snapshots_and_changes_for_other_tasks_are_refused()
    {
        using var t = new TestResources();
        t.CreateTask(@"C:\PPTK-Test\x\PPTKKeys.dll");
        string foreign = new JsonObject { ["tasks"] = new JsonArray(new JsonObject { ["path"] = "\\", ["name"] = "PPTKKeys Startup", ["xml"] = "<Task/>" }) }.ToJsonString();
        Assert.Throws<InvalidOperationException>(() => t.Env.ImportTasks(foreign));
        string otherTest = new JsonObject { ["tasks"] = new JsonArray(new JsonObject { ["path"] = "\\", ["name"] = "PPTK-Test-ffffffffffff", ["xml"] = "<Task/>" }) }.ToJsonString();
        Assert.Throws<InvalidOperationException>(() => t.Env.ImportTasks(otherTest));
        t.AssertProductionUnchanged();
    }

    [Fact]
    public void Setting_the_dll_of_a_missing_task_is_refused()
    {
        using var t = new TestResources();
        Assert.Throws<InvalidOperationException>(() => t.Env.SetHelperTaskDll(@"C:\x\PPTKKeys.dll"));
        Assert.Null(t.Env.GetHelperTaskDll());
        t.AssertProductionUnchanged();
    }

    [Fact]
    public void Registry_round_trip_uses_dword_and_sz_and_restores_exactly()
    {
        using var t = new TestResources();
        Assert.Null(t.Env.ExportRegistryKey(PendingAddinUpdate.AddinKey));
        t.Env.ImportRegistryKey(PendingAddinUpdate.AddinKey, "{\"Path\":\"PPTK.ppam\",\"AutoLoad\":1}");
        using (var k = Registry.CurrentUser.OpenSubKey(t.Config.MapKey(PendingAddinUpdate.AddinKey))!)
        {
            Assert.Equal(RegistryValueKind.DWord, k.GetValueKind("AutoLoad"));
            Assert.Equal(RegistryValueKind.String, k.GetValueKind("Path"));
        }
        var exported = JsonNode.Parse(t.Env.ExportRegistryKey(PendingAddinUpdate.AddinKey)!)!;
        Assert.Equal(1, (int)exported["AutoLoad"]!);
        Assert.Equal("PPTK.ppam", (string?)exported["Path"]);

        t.Env.ImportRegistryKey(PendingAddinUpdate.PptkKey, "{\"InstalledVersion\":\"1.0.0\",\"Stale\":\"x\"}");
        t.Env.ImportRegistryKey(PendingAddinUpdate.PptkKey, "{\"InstalledVersion\":\"1.0.1\"}");   // exact restore: Stale is gone
        Assert.Equal("{\"InstalledVersion\":\"1.0.1\"}", t.Env.ExportRegistryKey(PendingAddinUpdate.PptkKey));
        t.Env.ImportRegistryKey(PendingAddinUpdate.PptkKey, null);
        Assert.Null(t.Env.ExportRegistryKey(PendingAddinUpdate.PptkKey));
        t.AssertProductionUnchanged();
    }

    // The engine with the live task and registry (test names) and simulated files in a temp root.
    static TestResources LiveFor(PendingFixture f)
    {
        var t = new TestResources();
        t.CreateTask(f.Paths.HelperDll(f.Installed));
        t.Env.ImportRegistryKey(PendingAddinUpdate.PptkKey, new JsonObject { ["InstalledVersion"] = f.Installed }.ToJsonString());
        t.Env.ImportRegistryKey(PendingAddinUpdate.AddinKey, "{\"Path\":\"PPTK.ppam\",\"AutoLoad\":1}");
        return t;
    }

    [Fact]
    public void Engine_applies_a_pending_update_through_the_live_task_and_registry()
    {
        using var f = new PendingFixture(keysChanged: true);
        using var t = LiveFor(f);
        var r = new PendingAddinUpdate(f.Paths, t.Env).TryApply();
        Assert.True(r.Outcome == PendingUpdateOutcome.Applied, $"{r.Outcome}: {r.Detail}");
        Assert.Equal(f.Paths.HelperDll(f.Pending), t.Env.GetHelperTaskDll());
        Assert.Equal(f.Pending, (string?)JsonNode.Parse(t.Env.ExportRegistryKey(PendingAddinUpdate.PptkKey)!)!["InstalledVersion"]);
        Assert.Equal(PendingFixture.Sha(f.VersionAddin(f.Pending)), PendingFixture.Sha(f.Paths.AddinFile));
        Assert.False(t.TaskInfo().Enabled);
        t.AssertProductionUnchanged();
    }

    [Theory]
    [InlineData("task")]
    [InlineData("replaced")]
    [InlineData("registry")]
    public void Interrupted_live_update_is_undone_from_a_task_xml_journal_then_applied(string step)
    {
        using var f = new PendingFixture(keysChanged: true);
        using var t = LiveFor(f);
        var r = new PendingAddinUpdate(f.Paths, t.Env, hooks: new PendingUpdateTestHooks { AbandonAfter = s => s == step }).TryApply();
        Assert.True(r.Outcome == PendingUpdateOutcome.Interrupted, $"{r.Outcome}: {r.Detail}");
        Assert.Equal("task-xml", (string?)JsonNode.Parse(File.ReadAllText(f.Paths.JournalPath))!["tasksFormat"]);

        // The undo alone (a pass that defers after recovering): everything back to 1.0.0.
        t.Probe.Processes.Add(new PowerPointProcess(1, DateTime.UtcNow, true, "explorer.exe"));
        var deferred = new PendingAddinUpdate(f.Paths, t.Env).TryApply();
        Assert.True(deferred.Outcome == PendingUpdateOutcome.Deferred, $"{deferred.Outcome}: {deferred.Detail}");
        Assert.Equal(f.Paths.HelperDll(f.Installed), t.Env.GetHelperTaskDll());
        Assert.Equal(f.Installed, (string?)JsonNode.Parse(t.Env.ExportRegistryKey(PendingAddinUpdate.PptkKey)!)!["InstalledVersion"]);
        Assert.Equal(PendingFixture.Sha(f.VersionAddin(f.Installed)), PendingFixture.Sha(f.Paths.AddinFile));
        Assert.False(File.Exists(f.Paths.JournalPath));

        t.Probe.Processes.Clear();
        var applied = new PendingAddinUpdate(f.Paths, t.Env).TryApply();
        Assert.True(applied.Outcome == PendingUpdateOutcome.Applied, $"{applied.Outcome}: {applied.Detail}");
        Assert.Equal(f.Paths.HelperDll(f.Pending), t.Env.GetHelperTaskDll());
        t.AssertProductionUnchanged();
    }

    [Fact]
    public void Live_journal_is_refused_by_the_simulated_environment_and_vice_versa()
    {
        using var f = new PendingFixture();
        using var t = LiveFor(f);
        Assert.Equal(PendingUpdateOutcome.Interrupted, new PendingAddinUpdate(f.Paths, t.Env, hooks: new PendingUpdateTestHooks { AbandonAfter = s => s == "task" }).TryApply().Outcome);
        var r = new PendingAddinUpdate(f.Paths, f.Env).TryApply();   // a sim-json environment meets a task-xml journal
        Assert.True(r.Outcome == PendingUpdateOutcome.RestoreFailed, $"{r.Outcome}: {r.Detail}");
        Assert.Contains("task-xml", r.Detail);
        Assert.Equal(PendingUpdateOutcome.Applied, new PendingAddinUpdate(f.Paths, t.Env).TryApply().Outcome);   // its own environment recovers it
        t.AssertProductionUnchanged();
    }

    [Fact]
    public void Live_probe_only_reads_and_matches_the_session_process_list()
    {
        var before = Process.GetProcessesByName("POWERPNT").Select(p => p.Id).OrderBy(i => i).ToList();
        var seen = new LivePowerPointProbe().CurrentSession();
        var after = Process.GetProcessesByName("POWERPNT").Select(p => p.Id).OrderBy(i => i).ToList();
        Assert.Equal(before, after);
        int session = Process.GetCurrentProcess().SessionId;
        var expected = before.Where(id => { try { return Process.GetProcessById(id).SessionId == session; } catch (ArgumentException) { return false; } }).ToList();
        Assert.Equal(expected, seen.Select(p => p.ProcessId).OrderBy(i => i).ToList());
        Assert.All(seen, p => Assert.True(p.StartTimeUtc.Kind == DateTimeKind.Utc));
    }
}
