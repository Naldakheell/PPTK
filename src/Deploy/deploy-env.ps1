<#
  Environment layer shared by install.ps1 and uninstall.ps1 (dot-sourced; not run on its own).

  Three modes, chosen by the calling script's parameters:
    production  no -Root. The signed-in user's real profile folders, HKCU, Task Scheduler task
                "\PPTKKeys Startup", PowerPoint processes of this Windows session, and the running helper.
                Refused for SYSTEM and service accounts (PPTK is a per-user installation: Intune
                "Install behavior" must be User).
    sim         -Root <test root>. Everything simulated with files in the test root (registry.json,
                tasks.json, processes.json, helper-actions.log); nothing live is read or written.
    live-test   -Root <test root> -LiveTestId <12 hex>. Folders in the test root, but the REAL Task
                Scheduler and HKCU under test names only: task "\PPTK-Test-<id>", keys under
                HKCU\Software\PPTK-Test-<id>\. PowerPoint and the helper stay simulated
                (processes.json, helper-actions.log): the user's PowerPoint and helper are never touched.

  The caller defines: $Root, $LiveTestId, $Utf8, Log, Has, Read-Json, Write-Atomic, ConvertTo-JsonText,
  and the message prefix in $EnvPrefix. Refusals end the process ([Environment]::Exit): 'exit' in a dot-sourced
  file would only leave this file.
#>

$PptkKey = 'Software\PPTK'
$AddinKey = 'Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK'
$TaskPath = '\'
$ProductionTaskName = 'PPTKKeys Startup'

# The 64-bit Program Files folder, also when this runs in a 32-bit PowerShell.
$ProgramFiles64 = if ($env:ProgramW6432) { $env:ProgramW6432 } else { [Environment]::GetFolderPath('ProgramFiles') }

function Exit-Unsafe([string]$msg) { [Console]::Error.WriteLine("$EnvPrefix UnsafeRoot: $msg; nothing was changed."); [Environment]::Exit(2) }

if (-not $Root) {
    if ($LiveTestId) { Exit-Unsafe '-LiveTestId needs -Root' }
    $Mode = 'production'
    $me = [Security.Principal.WindowsIdentity]::GetCurrent()
    if ($me.User.Value -in 'S-1-5-18', 'S-1-5-19', 'S-1-5-20') {
        [Console]::Error.WriteLine("$EnvPrefix WrongContext: running as $($me.Name). PPTK is installed per user: run it in the signed-in user's context (Intune install behavior: User); nothing was changed."); [Environment]::Exit(1)
    }
    $Identity = [pscustomobject]@{ name = $me.Name; sid = $me.User.Value }
    $AppData = [Environment]::GetFolderPath('ApplicationData')
    $LocalAppData = [Environment]::GetFolderPath('LocalApplicationData')
    if (-not $AppData -or -not $LocalAppData) { [Console]::Error.WriteLine("$EnvPrefix WrongContext: this account has no user profile folders; nothing was changed."); [Environment]::Exit(1) }
    $TaskName = $ProductionTaskName
    $DotnetDirs = @((Join-Path $LocalAppData 'Microsoft\dotnet'), (Join-Path $ProgramFiles64 'dotnet'))
} else {
    try { $Root = (Resolve-Path -LiteralPath $Root -ErrorAction Stop).Path.TrimEnd('\') } catch { Exit-Unsafe "$Root does not exist" }
    if (-not (Test-Path -LiteralPath (Join-Path $Root 'PPTK-TEST-ROOT'))) { Exit-Unsafe "$Root is not a PPTK test root (PPTK-TEST-ROOT is required)" }
    # A test root never overlaps the real PPTK folders.
    foreach ($real in (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PPTK'), (Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'Microsoft\AddIns')) {
        $r = [IO.Path]::GetFullPath($real).TrimEnd('\')
        if ($Root -ieq $r -or $Root.StartsWith("$r\", [StringComparison]::OrdinalIgnoreCase) -or $r.StartsWith("$Root\", [StringComparison]::OrdinalIgnoreCase)) { Exit-Unsafe "$Root overlaps the real folder $r" }
    }
    $AppData = Join-Path $Root 'AppData'
    $LocalAppData = Join-Path $Root 'LocalAppData'
    if ($LiveTestId) {
        if ($LiveTestId -cnotmatch '^[0-9a-f]{12}$') { Exit-Unsafe '-LiveTestId must be 12 lowercase hex characters' }
        $Mode = 'live-test'
        $me = [Security.Principal.WindowsIdentity]::GetCurrent()
        $Identity = [pscustomobject]@{ name = $me.Name; sid = $me.User.Value }
        $TaskName = "PPTK-Test-$LiveTestId"
        $DotnetDirs = @((Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Microsoft\dotnet'), (Join-Path $ProgramFiles64 'dotnet'))
    } else {
        if (-not (Test-Path -LiteralPath (Join-Path $Root 'identity.json'))) { Exit-Unsafe "$Root has no identity.json" }
        $Mode = 'sim'
        $Identity = Read-Json (Join-Path $Root 'identity.json')
        $TaskName = $ProductionTaskName
        $DotnetDirs = @((Join-Path $LocalAppData 'Microsoft\dotnet'), (Join-Path $Root 'ProgramFiles\dotnet'))
    }
}
$TasksFormat = if ($Mode -eq 'sim') { 'sim-json' } else { 'task-xml' }

# ---------------------------------------------------------------- HKCU
function Get-MappedKey([string]$key) {
    if ($key -notin $PptkKey, $AddinKey) { throw "unexpected registry key '$key'" }
    if ($Mode -eq 'live-test') { return "Software\PPTK-Test-$LiveTestId\$key" }
    return $key
}
function Get-RegPath([string]$key) { return 'Registry::HKEY_CURRENT_USER\' + (Get-MappedKey $key) }

function Get-Reg([string]$key, [string]$name) {
    if ($Mode -eq 'sim') { $r = Read-Json (Join-Path $Root 'registry.json'); if ((Has $r $key) -and (Has $r.$key $name)) { return $r.$key.$name } else { return $null } }
    $p = Get-ItemProperty -LiteralPath (Get-RegPath $key) -Name $name -ErrorAction SilentlyContinue
    if ($p) { return $p.$name } else { return $null }
}
function Test-RegKey([string]$key) {
    if ($Mode -eq 'sim') { return Has (Read-Json (Join-Path $Root 'registry.json')) $key }
    return Test-Path -LiteralPath (Get-RegPath $key)
}
# Replaces the key with exactly these values (numbers REG_DWORD, everything else REG_SZ); $null deletes it.
function Set-RegKey([string]$key, $values) {
    if ($Mode -eq 'sim') {
        $r = Read-Json (Join-Path $Root 'registry.json'); if (-not $r) { $r = [pscustomobject]@{} }
        if ($null -eq $values) { if (Has $r $key) { $r.PSObject.Properties.Remove($key) } }
        else { $r | Add-Member -NotePropertyName $key -NotePropertyValue $values -Force }
        Write-Atomic (Join-Path $Root 'registry.json') (ConvertTo-JsonText $r)
        return
    }
    $path = Get-RegPath $key
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force }
    if ($null -eq $values) { return }
    New-Item -Path $path -Force | Out-Null
    $names = if ($values -is [Collections.IDictionary]) { @($values.Keys) } else { @($values.PSObject.Properties | ForEach-Object { $_.Name }) }
    foreach ($n in $names) {
        $v = if ($values -is [Collections.IDictionary]) { $values[$n] } else { $values.$n }
        if ($v -is [int] -or $v -is [long]) { New-ItemProperty -LiteralPath $path -Name $n -Value ([int]$v) -PropertyType DWord -Force | Out-Null }
        else { New-ItemProperty -LiteralPath $path -Name $n -Value ([string]$v) -PropertyType String -Force | Out-Null }
    }
}
# The key's values as JSON object text (journal), or $null when the key does not exist.
function Export-RegKey([string]$key) {
    if ($Mode -eq 'sim') { $r = Read-Json (Join-Path $Root 'registry.json'); if (Has $r $key) { return ConvertTo-JsonText $r.$key } else { return $null } }
    $path = Get-RegPath $key
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $k = Get-Item -LiteralPath $path
    $o = [ordered]@{}
    foreach ($n in @($k.GetValueNames() | Sort-Object)) {
        $v = $k.GetValue($n, $null, 'DoNotExpandEnvironmentNames')
        $o[$n] = if ($k.GetValueKind($n) -eq 'DWord') { [int]$v } else { [string]$v }
    }
    return ConvertTo-JsonText ([pscustomobject]$o)
}

# ---------------------------------------------------------------- Task Scheduler
# The PPTKKeys tasks, in the sim schema (detect.ps1): production = "\PPTKKeys Startup" and every task whose
# action runs PPTKKeys.dll; live-test = only this run's test task.
function Test-RunsPptkKeys($t) { return @(@($t.actions) | Where-Object { "$($_.execute) $($_.arguments)" -match 'PPTKKeys\.dll' }).Count -gt 0 }
function Get-Tasks {
    if ($Mode -eq 'sim') { return @(Read-Json (Join-Path $Root 'tasks.json') | Where-Object { $_ }) }
    $out = @()
    foreach ($t in @(Get-ScheduledTask -ErrorAction SilentlyContinue)) {
        $actions = @($t.Actions | Where-Object { $_.CimClass.CimClassName -eq 'MSFT_TaskExecAction' } | ForEach-Object { [pscustomobject]@{ execute = $_.Execute; arguments = $_.Arguments } })
        $o = [pscustomobject]@{
            name = $t.TaskName; path = $t.TaskPath; enabled = ($t.State -ne 'Disabled')
            userId = [string]$t.Principal.UserId; logonType = [string]$t.Principal.LogonType; runLevel = [string]$t.Principal.RunLevel
            triggers = @($t.Triggers | ForEach-Object { [pscustomobject]@{ type = $(if ($_.CimClass.CimClassName -eq 'MSFT_TaskLogonTrigger') { 'Logon' } else { $_.CimClass.CimClassName }); userId = [string]$_.UserId; enabled = [bool]$_.Enabled } })
            actions = $actions
        }
        if ($Mode -eq 'live-test') { if ($o.name -eq $TaskName -and $o.path -eq $TaskPath) { $out += $o } }
        elseif (($o.name -eq $TaskName -and $o.path -eq $TaskPath) -or (Test-RunsPptkKeys $o)) { $out += $o }
    }
    return $out
}

function New-TaskEntry([string]$dotnet, [string]$dll) {
    return [ordered]@{
        name = $TaskName; path = $TaskPath; enabled = $true
        userId = $Identity.name; logonType = 'Interactive'; runLevel = 'Limited'
        triggers = @([ordered]@{ type = 'Logon'; userId = $Identity.name; enabled = $true })
        actions = @([ordered]@{ execute = $dotnet; arguments = 'exec "' + $dll + '"'; workingDirectory = (Split-Path -Parent $dll) })
        settings = [ordered]@{ multipleInstances = 'IgnoreNew'; executionTimeLimit = 'PT0S' }
    }
}

# Task XML: current user, interactive token (no password, no elevation), least privilege, logon trigger,
# runs on battery, no time limit, one instance.
function New-TaskXml([string]$dotnet, [string]$dll) {
    $x = { param($s) [Security.SecurityElement]::Escape($s) }
    return @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Starts the PPTKKeys keyboard helper when the current user signs in.</Description>
  </RegistrationInfo>
  <Principals>
    <Principal id="Author">
      <UserId>$(& $x $Identity.sid)</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <StartWhenAvailable>true</StartWhenAvailable>
    <Enabled>true</Enabled>
  </Settings>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>$(& $x $Identity.sid)</UserId>
    </LogonTrigger>
  </Triggers>
  <Actions Context="Author">
    <Exec>
      <Command>$(& $x $dotnet)</Command>
      <Arguments>$(& $x ('exec "' + $dll + '"'))</Arguments>
      <WorkingDirectory>$(& $x (Split-Path -Parent $dll))</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@
}

function Register-TaskXml([string]$name, [string]$xml) {
    if ($Mode -eq 'live-test' -and $name -ne $TaskName) { throw "live-test refuses task '$name'" }
    if ($Mode -eq 'production' -and $name -ne $ProductionTaskName) { throw "refusing to register task '$name'" }
    Register-ScheduledTask -Xml $xml -TaskName $name -TaskPath $TaskPath -Force -ErrorAction Stop | Out-Null
}

# Create or replace "\<TaskName>" so it runs the given helper.
function Set-MainTask([string]$dotnet, [string]$dll) {
    if ($Mode -eq 'sim') {
        Set-SimTasks (@(Get-Tasks | Where-Object { -not ($_.name -eq $TaskName -and $_.path -eq $TaskPath) }) + @(New-TaskEntry $dotnet $dll))
        return
    }
    Register-TaskXml $TaskName (New-TaskXml $dotnet $dll)
}
function Set-SimTasks($tasks) { Write-Atomic (Join-Path $Root 'tasks.json') (ConvertTo-JsonText @($tasks)) }

function Remove-MainTask {
    if ($Mode -eq 'sim') { Set-SimTasks @(Get-Tasks | Where-Object { -not ($_.name -eq $TaskName -and $_.path -eq $TaskPath) }); return }
    if (Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop }
}

# Remove only PPTKKeys tasks (the main task and tasks running PPTKKeys.dll); every other task is kept.
function Remove-PptkTasks {
    if ($Mode -eq 'sim') { Set-SimTasks @(Get-Tasks | Where-Object { -not ($_.name -eq $TaskName -and $_.path -eq $TaskPath) -and -not (Test-RunsPptkKeys $_) }); return }
    foreach ($t in @(Get-Tasks)) { Unregister-ScheduledTask -TaskName $t.name -TaskPath $t.path -Confirm:$false -ErrorAction Stop }
}

# Journal snapshot of the PPTKKeys task (src\Deploy\JOURNAL.md): sim-json = tasks.json text; task-xml =
# {"tasks":[{path,name,xml}]} with only the main task. $null = there was no such task.
function Export-Tasks {
    if ($Mode -eq 'sim') { $p = Join-Path $Root 'tasks.json'; if (Test-Path -LiteralPath $p) { return [IO.File]::ReadAllText($p) } else { return $null } }
    $t = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
    if (-not $t) { return $null }
    $xml = Export-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath
    return ConvertTo-Json -Depth 5 -InputObject ([ordered]@{ tasks = @([ordered]@{ path = $TaskPath; name = $TaskName; xml = [string]$xml }) })
}
function Import-Tasks($snapshot) {
    if ($Mode -eq 'sim') { if ($null -ne $snapshot) { Write-Atomic (Join-Path $Root 'tasks.json') $snapshot } else { Remove-PptkTasks }; return }
    if ($null -eq $snapshot) { Remove-MainTask; return }   # only the task this operation may have created
    $entries = @(($snapshot | ConvertFrom-Json).tasks)
    if (-not $entries.Count) { throw 'task-xml snapshot has no tasks' }
    foreach ($e in $entries) {
        if ($e.path -ne $TaskPath -or $e.name -ne $TaskName) { throw "task-xml snapshot names '$($e.path)$($e.name)', not '$TaskPath$TaskName'" }
        Register-TaskXml $e.name ([string]$e.xml)
    }
}
function Assert-JournalFormat($j) {
    if ((Has $j 'journalVersion') -and $j.journalVersion -ne 1) { throw "unknown journalVersion $($j.journalVersion)" }
    $fmt = if ((Has $j 'tasksFormat') -and $j.tasksFormat) { $j.tasksFormat } else { 'sim-json' }
    if ($fmt -ne $TasksFormat) { throw "the journal's tasks are in '$fmt' format; this environment uses '$TasksFormat'" }
}

# ---------------------------------------------------------------- architecture and the private .NET runtime
$RuntimeRids = @{ 'win-x64' = 0x8664; 'win-arm64' = 0xAA64 }   # runtime identifier -> PE machine of its dotnet.exe

# The machine's own architecture, not this process's: on ARM64 an x64 or x86 PowerShell runs emulated and
# reports the emulated architecture. Test roots may set it in <Root>\architecture.txt.
function Get-NativeRid {
    if ($Mode -ne 'production') { $p = Join-Path $Root 'architecture.txt'; if (Test-Path -LiteralPath $p) { return ([IO.File]::ReadAllText($p)).Trim() } }
    $a = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | ForEach-Object { $_.Architecture } | Select-Object -First 1)
    if ($a.Count) { switch ([int]$a[0]) { 9 { return 'win-x64' } 12 { return 'win-arm64' } } }
    $n = (Get-ItemProperty -LiteralPath 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment' -Name PROCESSOR_ARCHITECTURE -ErrorAction SilentlyContinue).PROCESSOR_ARCHITECTURE
    switch ($n) { 'AMD64' { return 'win-x64' } 'ARM64' { return 'win-arm64' } }
    return "unsupported (processor architecture $(if ($a.Count) { $a[0] } else { '?' }) / $n)"
}

function Get-PeMachine([string]$path) {
    $fs = [IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
    try {
        $b = New-Object byte[] 64; [void]$fs.Read($b, 0, 64)
        $fs.Position = [BitConverter]::ToInt32($b, 0x3c) + 4
        $m = New-Object byte[] 2; [void]$fs.Read($m, 0, 2)
        return [int][BitConverter]::ToUInt16($m, 0)
    } finally { $fs.Dispose() }
}

function Test-MicrosoftSigned([string]$path) {
    $s = Get-AuthenticodeSignature -LiteralPath $path
    return ($s.Status -eq 'Valid') -and ($s.SignerCertificate.Subject -match '(^|,\s*)O=Microsoft Corporation(,|$)')
}

# A runtime folder (the content of an official dotnet-runtime-<v>-<rid>.zip) ready to run PPTKKeys: every
# .exe/.dll signed by Microsoft, dotnet.exe built for this architecture, and the framework PPTKKeys needs.
# Throws with the reason.
function Assert-Runtime([string]$dir, [string]$rid, [string]$framework, [version]$minimum) {
    $host_ = Join-Path $dir 'dotnet.exe'
    if (-not (Test-Path -LiteralPath $host_)) { throw "the $rid runtime has no dotnet.exe" }
    if ((Get-PeMachine $host_) -ne $RuntimeRids[$rid]) { throw ('the {0} runtime''s dotnet.exe is built for machine 0x{1:X4}, not {0}' -f $rid, (Get-PeMachine $host_)) }
    $dir = (Get-Item -LiteralPath $dir).FullName.TrimEnd('\')
    $unsigned = @(Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.Extension -in '.exe', '.dll' } | Where-Object { -not (Test-MicrosoftSigned $_.FullName) } | ForEach-Object { $_.FullName.Substring($dir.Length + 1) })
    if ($unsigned.Count) { throw ("files in the $rid runtime are not validly signed by Microsoft: " + (($unsigned | Select-Object -First 5) -join ', ')) }
    $fw = Join-Path $dir "shared\$framework"
    $ok = @(if (Test-Path -LiteralPath $fw) { Get-ChildItem -LiteralPath $fw -Directory | Where-Object { $_.Name -match '^\d+\.\d+\.\d+$' } | Where-Object { $v = [version]$_.Name; $v.Major -eq $minimum.Major -and $v -ge $minimum } })
    if (-not $ok.Count) { throw "the $rid runtime does not contain $framework $minimum (same major version)" }
}

# SHA-256 over "<path relative to the runtime folder, forward slashes> <sha256>" lines, sorted ordinally, one per
# file. Equal for the package's runtime\<rid>\ and an installed versions\<v>\runtime\ (src\Deploy\build-package.ps1).
function Get-RuntimeFingerprint([string]$dir) {
    $dir = (Get-Item -LiteralPath $dir).FullName.TrimEnd('\')   # the long form: Get-ChildItem returns long names under an 8.3 path
    $lines = [Collections.Generic.List[string]]::new()
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Recurse -File -Force)) { $lines.Add($f.FullName.Substring($dir.Length + 1).Replace('\', '/') + ' ' + (Get-Sha256 $f.FullName)) }
    $lines.Sort([StringComparer]::Ordinal)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($lines -join "`n") + "`n"))).Replace('-', '')).ToLowerInvariant() } finally { $sha.Dispose() }
}

# ---------------------------------------------------------------- PowerPoint and the helper
function Test-PowerPointRunning {
    if ($Mode -ne 'production') { return @(@(Read-Json (Join-Path $Root 'processes.json')) | Where-Object { $_ -ieq 'POWERPNT' }).Count -gt 0 }
    $session = (Get-Process -Id $PID).SessionId
    return @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $session }).Count -gt 0
}

# Running PPTKKeys helpers of this session: each has a message-only window (title "PPTKKeys", which older builds register as "P") of class
# "PPTKKeys.Message.<pid>". Read-only.
function Get-HelperWindows {
    if (-not ('PPTKDeploy.HelperWindows' -as [type])) {
        Add-Type -Namespace PPTKDeploy -Name HelperWindows -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
public static extern System.IntPtr FindWindowEx(System.IntPtr parent, System.IntPtr after, string cls, string title);
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
public static extern int GetClassName(System.IntPtr hwnd, System.Text.StringBuilder name, int max);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hwnd, out uint pid);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool PostMessage(System.IntPtr hwnd, uint msg, System.IntPtr w, System.IntPtr l);
'@
    }
    $found = @(); $h = [IntPtr]::Zero
    while ($true) {
        $h = [PPTKDeploy.HelperWindows]::FindWindowEx([IntPtr](-3), $h, [NullString]::Value, [NullString]::Value)   # every message-only window; matched by class below
        if ($h -eq [IntPtr]::Zero) { break }
        $sb = New-Object Text.StringBuilder 128; [void][PPTKDeploy.HelperWindows]::GetClassName($h, $sb, 128)
        $hp = [uint32]0; [void][PPTKDeploy.HelperWindows]::GetWindowThreadProcessId($h, [ref]$hp)
        if ($sb.ToString() -eq "PPTKKeys.Message.$hp") { $found += [pscustomobject]@{ Hwnd = $h; Pid = [int]$hp } }
    }
    return $found
}

# Ask every running helper to close (WM_CLOSE: it removes its hook and releases its mutex) and wait up to
# $timeoutSeconds. Never terminates a process. $true when no helper is left.
function Stop-Helper([int]$timeoutSeconds = 15) {
    try { $helpers = @(Get-HelperWindows) } catch { Log "helper: cannot look for the running helper ($($_.Exception.Message))"; return $false }
    if (-not $helpers.Count) { Log 'helper: none running'; return $true }
    # Process objects taken before asking, so a reused pid is never mistaken for the helper.
    $procs = @($helpers | ForEach-Object { Get-Process -Id $_.Pid -ErrorAction SilentlyContinue } | Where-Object { $_ })
    foreach ($w in $helpers) { [void][PPTKDeploy.HelperWindows]::PostMessage($w.Hwnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero); Log "helper: asked pid $($w.Pid) to close" }
    $deadline = [DateTime]::UtcNow.AddSeconds($timeoutSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        if (-not @($procs | Where-Object { -not $_.HasExited }).Count) { Log 'helper: closed'; return $true }
        Start-Sleep -Milliseconds 250
    }
    Log ('helper: still running after ' + $timeoutSeconds + ' s: pid ' + (($procs | Where-Object { -not $_.HasExited } | ForEach-Object { $_.Id }) -join ', '))
    return $false
}

# Production undo: the helper may already run from the folder the undo removes; ask it to close first.
function Stop-HelperForUndo { if ($Mode -eq 'production') { [void](Stop-Helper) } }
function Remove-CreatedVersionDir([string]$rel) {
    if (-not $rel) { return }
    $dir = Join-Path $PptkDir $rel
    if (-not (Test-Path -LiteralPath $dir)) { return }
    try { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction Stop }
    catch { Log "$rel could not be removed now ($($_.Exception.Message)); it is not used by the restored installation and is removed by a later run" }
}

# stop / start / restart of the helper task. Simulated (helper-actions.log) outside production.
# Returns $false only when a stop was asked and the helper did not close.
function Helper-Action([string]$what) {
    if ($Mode -ne 'production') { [IO.File]::AppendAllText((Join-Path $Root 'helper-actions.log'), "$what`r`n", $Utf8); Log "simulated helper task: $what"; return $true }
    $ok = $true
    if ($what -in 'stop', 'restart (restored)') { $ok = Stop-Helper }
    if ($what -in 'start', 'restart (restored)') {
        if (Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue) {
            try { Start-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction Stop | Out-Null; Log 'helper: task started' } catch { Log "helper: task not started ($($_.Exception.Message)); it starts at the next sign-in" }
        }
    }
    return $ok
}
