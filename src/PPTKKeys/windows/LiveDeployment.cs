using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace PPTKKeys;

// Phase 2B-3a: the live implementations of IDeploymentEnvironment and IPowerPointProbe.
// Production use is switched off: only a test configuration (PPTK-Test-* names) can be built, and it
// refuses every production task, registry key and path. Not connected to the helper.

public static class DeploymentFlags
{
    /// <summary>Pending updates, the update notice and the handoff for the installed helper (a helper running from
    /// LocalAppData\PPTK\versions\&lt;v&gt;\PPTKKeys). On since the installer writes and restores "task-xml"
    /// journals (src\Deploy\JOURNAL.md).</summary>
    public static readonly bool ProductionUpdatesEnabled = true;

    /// <summary>Terminating an unresponsive new helper during a failed handoff. Off: not validated on a real
    /// device yet. Without it a failed recovery leaves the hook with the new helper (which owns the mutex).</summary>
    public static readonly bool ForcedRecoveryEnabled = false;
}

/// <summary>Where the live environment acts. A test configuration maps the logical keys and task to
/// PPTK-Test-* names; production names are refused.</summary>
public sealed class LiveDeploymentConfig
{
    public const string ProductionTaskName = "PPTKKeys Startup";
    static readonly string[] ProductionKeys = [PendingAddinUpdate.PptkKey, PendingAddinUpdate.AddinKey];
    static readonly Regex TestName = new(@"^PPTK-Test-[0-9a-f]{8,32}$", RegexOptions.CultureInvariant);

    LiveDeploymentConfig(string taskName, string? registryRoot, bool production)
    {
        TaskName = taskName; RegistryRoot = registryRoot; IsProduction = production;
    }

    /// <summary>Task in the root folder ("\").</summary>
    public string TaskName { get; }
    /// <summary>HKCU sub key the logical keys are mapped under; null in production (identity mapping).</summary>
    public string? RegistryRoot { get; }
    public bool IsProduction { get; }

    /// <summary>Test configuration: task "\PPTK-Test-&lt;hex&gt;" and keys under HKCU\Software\PPTK-Test-&lt;hex&gt;\.</summary>
    public static LiveDeploymentConfig ForTest(string taskName, string registryRoot)
    {
        if (!TestName.IsMatch(taskName)) throw new ArgumentException($"test task name must be PPTK-Test-<hex>, not '{taskName}'");
        var parts = registryRoot.Split('\\');
        if (parts.Length != 2 || !parts[0].Equals("Software", StringComparison.OrdinalIgnoreCase) || !TestName.IsMatch(parts[1]))
            throw new ArgumentException($@"test registry root must be Software\PPTK-Test-<hex>, not '{registryRoot}'");
        return new LiveDeploymentConfig(taskName, registryRoot, false);
    }

    /// <summary>The production configuration. Refused while DeploymentFlags.ProductionUpdatesEnabled is off.</summary>
    public static LiveDeploymentConfig Production()
    {
        if (!DeploymentFlags.ProductionUpdatesEnabled) throw new InvalidOperationException("production update activation is disabled in this build");
        return new LiveDeploymentConfig(ProductionTaskName, null, true);
    }

    /// <summary>The HKCU path a logical key maps to. A test configuration never maps to a production key.</summary>
    public string MapKey(string logicalKey)
    {
        if (!ProductionKeys.Contains(logicalKey, StringComparer.OrdinalIgnoreCase)) throw new ArgumentException($"unexpected key '{logicalKey}'");
        if (IsProduction) return logicalKey;
        string mapped = RegistryRoot + "\\" + logicalKey;
        if (ProductionKeys.Any(k => mapped.Equals(k, StringComparison.OrdinalIgnoreCase) || mapped.StartsWith(k + "\\", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException($"test key '{mapped}' would touch a production key");
        return mapped;
    }

    /// <summary>Throws unless the task name is this configuration's own (and, in test mode, a test name).</summary>
    public void CheckTask(string path, string name)
    {
        if (path != "\\" || name != TaskName) throw new InvalidOperationException($"task '{path}{name}' is not '\\{TaskName}'");
        if (!IsProduction && (name.Equals(ProductionTaskName, StringComparison.OrdinalIgnoreCase) || !TestName.IsMatch(name)))
            throw new InvalidOperationException($"test configuration refuses task '{name}'");
    }
}

[SupportedOSPlatform("windows")]
public sealed class LiveDeploymentEnvironment : IDeploymentEnvironment
{
    const int TaskCreateOrUpdate = 6, TaskUpdate = 4, TaskLogonInteractiveToken = 3, TaskActionExec = 0;
    static readonly Regex ExecArgs = new("^\\s*exec\\s+\"([^\"]+)\"\\s*$", RegexOptions.CultureInvariant);

    readonly LiveDeploymentConfig config;
    readonly IPowerPointProbe probe;

    public LiveDeploymentEnvironment(LiveDeploymentConfig config, IPowerPointProbe probe)
    {
        this.config = config;
        this.probe = probe;
        config.CheckTask("\\", config.TaskName);
    }

    public string TasksFormat => "task-xml";

    public bool IsPowerPointRunning() => probe.CurrentSession().Count > 0;

    // ------------------------------------------------------------ tasks (Task Scheduler COM, root folder only)
    static dynamic Connect()
    {
        dynamic service = Activator.CreateInstance(Type.GetTypeFromProgID("Schedule.Service", throwOnError: true)!)!;
        service.Connect();
        return service;
    }

    dynamic? FindTask(dynamic folder)
    {
        try { return folder.GetTask(config.TaskName); }
        catch (Exception ex) when (IsNotFound(ex)) { return null; }
    }

    // Through dynamic COM interop "not found" arrives as FileNotFoundException/DirectoryNotFoundException or COMException.
    internal static bool IsNotFound(Exception ex) => (uint)ex.HResult is 0x80070002 or 0x80070003;

    /// <summary>The configured task's principal must be the current user (SID, full name or bare name).</summary>
    static void CheckOwner(dynamic definition)
    {
        string user = (string)(definition.Principal.UserId ?? "");
        using var me = WindowsIdentity.GetCurrent();
        string bare = me.Name.Split('\\')[^1];
        bool mine = user.Equals(me.User?.Value, StringComparison.OrdinalIgnoreCase) || user.Equals(me.Name, StringComparison.OrdinalIgnoreCase)
            || (!user.Contains('\\') && user.Equals(bare, StringComparison.OrdinalIgnoreCase));
        if (!mine) throw new InvalidOperationException($"the task belongs to '{user}', not the current user");
    }

    public string? ExportTasks()
    {
        dynamic folder = Connect().GetFolder("\\");
        dynamic? task = FindTask(folder);
        if (task is null) return null;
        var tasks = new JsonArray(new JsonObject { ["path"] = "\\", ["name"] = config.TaskName, ["xml"] = (string)task.Xml });
        return new JsonObject { ["tasks"] = tasks }.ToJsonString();
    }

    public void ImportTasks(string? snapshot)
    {
        dynamic folder = Connect().GetFolder("\\");
        if (snapshot is null)
        {
            if (FindTask(folder) is not null) folder.DeleteTask(config.TaskName, 0);   // only the configured task
            return;
        }
        var entries = (JsonNode.Parse(snapshot)?["tasks"] as JsonArray) ?? throw new InvalidOperationException("task-xml snapshot has no tasks array");
        foreach (var e in entries)
        {
            string path = (string?)e?["path"] ?? "", name = (string?)e?["name"] ?? "", xml = (string?)e?["xml"] ?? "";
            config.CheckTask(path, name);   // a snapshot can restore only the configured task
            folder.RegisterTask(name, xml, TaskCreateOrUpdate, null, null, TaskLogonInteractiveToken, null);
        }
    }

    public string? GetHelperTaskDll()
    {
        dynamic? task = FindTask(Connect().GetFolder("\\"));
        if (task is null) return null;
        dynamic actions = task.Definition.Actions;
        if ((int)actions.Count != 1) return null;
        dynamic action = actions.Item(1);
        if ((int)action.Type != TaskActionExec) return null;
        var m = ExecArgs.Match((string)(action.Arguments ?? ""));
        return m.Success ? m.Groups[1].Value : null;
    }

    public string? GetHelperTaskHost()
    {
        dynamic? task = FindTask(Connect().GetFolder("\\"));
        if (task is null) return null;
        dynamic actions = task.Definition.Actions;
        if ((int)actions.Count != 1) return null;
        dynamic action = actions.Item(1);
        return (int)action.Type == TaskActionExec ? ((string?)action.Path)?.Trim('"') : null;
    }

    public void SetHelperTaskDll(string dll, string? host = null)
    {
        dynamic folder = Connect().GetFolder("\\");
        dynamic task = FindTask(folder) ?? throw new InvalidOperationException($"task '\\{config.TaskName}' does not exist");
        dynamic definition = task.Definition;
        CheckOwner(definition);
        dynamic actions = definition.Actions;
        if ((int)actions.Count != 1 || (int)actions.Item(1).Type != TaskActionExec) throw new InvalidOperationException("the task does not have exactly one program action");
        dynamic action = actions.Item(1);
        action.Arguments = $"exec \"{dll}\"";
        action.WorkingDirectory = Path.GetDirectoryName(dll);
        if (host is not null) action.Path = host;
        // TASK_UPDATE with the interactive token: no password, no elevation; everything else is kept.
        folder.RegisterTaskDefinition(config.TaskName, definition, TaskUpdate, null, null, TaskLogonInteractiveToken, null);
    }

    // ------------------------------------------------------------ HKCU (JSON numbers = REG_DWORD, strings = REG_SZ)
    public string? ExportRegistryKey(string key)
    {
        using var k = Registry.CurrentUser.OpenSubKey(config.MapKey(key));
        if (k is null) return null;
        var o = new JsonObject();
        foreach (var name in k.GetValueNames().OrderBy(n => n, StringComparer.Ordinal))
        {
            object? v = k.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            o[name] = k.GetValueKind(name) == RegistryValueKind.DWord ? JsonValue.Create((int)v!) : JsonValue.Create(Convert.ToString(v));
        }
        return o.ToJsonString();
    }

    public void ImportRegistryKey(string key, string? json)
    {
        string path = config.MapKey(key);
        if (json is null) { Registry.CurrentUser.DeleteSubKeyTree(path, throwOnMissingSubKey: false); return; }
        var values = JsonNode.Parse(json) as JsonObject ?? throw new InvalidOperationException("registry snapshot is not a JSON object");
        Registry.CurrentUser.DeleteSubKeyTree(path, throwOnMissingSubKey: false);   // exact restore: no stale values
        using var k = Registry.CurrentUser.CreateSubKey(path, writable: true);
        foreach (var (name, value) in values)
        {
            if (value is JsonValue jv && jv.GetValueKind() == JsonValueKind.Number) k.SetValue(name, jv.GetValue<int>(), RegistryValueKind.DWord);
            else k.SetValue(name, value?.GetValue<string>() ?? "", RegistryValueKind.String);
        }
    }
}

/// <summary>Read-only view of POWERPNT.EXE in the current Windows session: pid, start time, whether it has a
/// visible PPTFrameClass window, parent process name. No COM, no process control.</summary>
[SupportedOSPlatform("windows")]
public sealed partial class LivePowerPointProbe : IPowerPointProbe
{
    public IReadOnlyList<PowerPointProcess> CurrentSession()
    {
        int session = Process.GetCurrentProcess().SessionId;
        var processes = Process.GetProcessesByName("POWERPNT").Where(p => SafeSession(p) == session).ToList();
        if (processes.Count == 0) return [];
        var visible = VisibleFramePids();
        var parents = ParentIds();
        var result = new List<PowerPointProcess>();
        foreach (var p in processes)
        {
            using (p)
            {
                DateTime started; try { started = p.StartTime.ToUniversalTime(); } catch (Exception e) when (e is InvalidOperationException or System.ComponentModel.Win32Exception) { continue; }
                string? parent = parents.TryGetValue(p.Id, out int ppid) ? ProcessName(ppid) : null;
                result.Add(new PowerPointProcess(p.Id, started, visible.Contains(p.Id), parent));
            }
        }
        return result;
    }

    static int SafeSession(Process p) { try { return p.SessionId; } catch (InvalidOperationException) { return -1; } }

    static string? ProcessName(int pid)
    {
        try { using var p = Process.GetProcessById(pid); return p.ProcessName + ".exe"; }
        catch (ArgumentException) { return null; }
    }

    static HashSet<int> VisibleFramePids()
    {
        var pids = new HashSet<int>();
        var name = new StringBuilder(64);
        EnumWindows((hwnd, _) =>
        {
            if (!IsWindowVisible(hwnd)) return true;
            name.Clear();
            if (GetClassName(hwnd, name, name.Capacity) > 0 && name.ToString() == "PPTFrameClass")
            {
                GetWindowThreadProcessId(hwnd, out uint pid);
                pids.Add((int)pid);
            }
            return true;
        }, IntPtr.Zero);
        return pids;
    }

    // Parent process ids from a Toolhelp snapshot (read-only).
    static Dictionary<int, int> ParentIds()
    {
        var map = new Dictionary<int, int>();
        IntPtr snap = CreateToolhelp32Snapshot(0x2 /* TH32CS_SNAPPROCESS */, 0);
        if (snap == new IntPtr(-1)) return map;
        try
        {
            var entry = new ProcessEntry32 { dwSize = (uint)Marshal.SizeOf<ProcessEntry32>() };
            for (bool ok = Process32FirstW(snap, ref entry); ok; ok = Process32NextW(snap, ref entry)) map[(int)entry.th32ProcessID] = (int)entry.th32ParentProcessID;
        }
        finally { CloseHandle(snap); }
        return map;
    }

    delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder name, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint pid);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern bool Process32FirstW(IntPtr snap, ref ProcessEntry32 entry);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern bool Process32NextW(IntPtr snap, ref ProcessEntry32 entry);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct ProcessEntry32
    {
        public uint dwSize, cntUsage, th32ProcessID; public IntPtr th32DefaultHeapID; public uint th32ModuleID, cntThreads, th32ParentProcessID; public int pcPriClassBase; public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szExeFile;
    }
}
