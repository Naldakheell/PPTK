<#
.SYNOPSIS
  Live-adapter tests for install.ps1 / uninstall.ps1 / detect.ps1 (production code paths for Task Scheduler,
  HKCU and the task-xml journal), using test names only.

.DESCRIPTION
  Runs the scripts with -Root <temp test root> -LiveTestId <random 12 hex>: folders live in the test root,
  the REAL Task Scheduler and HKCU are used under "\PPTK-Test-<id>" and HKCU\Software\PPTK-Test-<id>\ only.
  PowerPoint and the helper are simulated (processes.json, helper-actions.log): the user's PowerPoint and
  helper are never touched. The production task, keys, add-in, PPTK folder and helper process are
  snapshotted before and after and must be unchanged. Everything the run created is removed in finally.
  The helper-window discovery used by production mode is checked read-only (no message is sent).
  Exit code 0 when every case passes.
#>
param([string]$RepoRoot = '')

$ErrorActionPreference = 'Stop'
if (-not $RepoRoot) { $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path }
$Install = Join-Path $RepoRoot 'src\Deploy\install.ps1'
$Uninstall = Join-Path $RepoRoot 'src\Deploy\uninstall.ps1'
$Detect = Join-Path $RepoRoot 'src\Deploy\detect.ps1'
$Id = [guid]::NewGuid().ToString('N').Substring(0, 12)
$TestTask = "PPTK-Test-$Id"
$TestKey = "HKCU:\Software\PPTK-Test-$Id"
$Base = Join-Path ([IO.Path]::GetTempPath()) ("PPTK-live-tests-$Id")
$Utf8 = New-Object Text.UTF8Encoding($false)
$results = New-Object System.Collections.ArrayList

function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) {
    [void]$results.Add([pscustomobject]@{ Id = $id; Ok = $ok })
    Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail)
}
function Sha([string]$p) { (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLowerInvariant() }
function Write-Json([string]$p, $o) { [IO.File]::WriteAllText($p, (ConvertTo-Json -InputObject $o -Depth 10), $Utf8) }

function New-FakePackage([string]$name, [string]$version, [string]$ppam, [string]$dll) {
    $dir = Join-Path $Base "pkg-$name"
    New-Item -ItemType Directory -Force -Path (Join-Path $dir 'payload\PPTKKeys') | Out-Null
    $content = [ordered]@{
        'payload/PPTK.ppam' = "add-in $ppam"
        'payload/PPTKKeys/PPTKKeys.dll' = "helper $dll"
        'payload/PPTKKeys/PPTKKeys.deps.json' = '{"runtimeTarget":{"name":".NETCoreApp,Version=v10.0"}}'
        'payload/PPTKKeys/PPTKKeys.runtimeconfig.json' = '{"runtimeOptions":{"tfm":"net10.0","framework":{"name":"Microsoft.NETCore.App","version":"10.0.0"}}}'
    }
    $files = @()
    foreach ($k in $content.Keys) { $p = Join-Path $dir ($k.Replace('/', '\')); [IO.File]::WriteAllText($p, $content[$k], $Utf8); $files += [ordered]@{ path = $k; size = (Get-Item -LiteralPath $p).Length; sha256 = (Sha $p) } }
    Write-Json (Join-Path $dir 'version.json') ([ordered]@{ schemaVersion = 1; packageName = 'PPTK'; packageVersion = $version; files = $files })
    return $dir
}

function Invoke-Ps([string]$script, [string[]]$arguments) {
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = $script:PsHost
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $script + '" ' + (($arguments | ForEach-Object { if ($_ -match '^-') { $_ } else { '"' + $_ + '"' } }) -join ' ')
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8; $psi.StandardErrorEncoding = [Text.Encoding]::UTF8
    $p = [Diagnostics.Process]::Start($psi); $o = $p.StandardOutput.ReadToEnd(); $e = $p.StandardError.ReadToEnd(); $p.WaitForExit()
    return [pscustomobject]@{ Code = $p.ExitCode; Out = $o.Trim(); Err = $e.Trim(); Text = ($o + ' ' + $e).Trim() }
}
$script:PsHost = 'powershell.exe'
$L = @('-Root', $Base, '-LiveTestId', $Id)
function Install([string]$pkg, [string[]]$extra = @()) { Invoke-Ps $Install (@('-PackagePath', $pkg) + $L + $extra) }
function Rollback { Invoke-Ps $Install (@('-Rollback') + $L) }
function Uninstall([string[]]$extra = @()) { Invoke-Ps $Uninstall ($L + $extra) }
function Detect([string]$pkg) { Invoke-Ps $Detect (@('-ManifestPath', (Join-Path $pkg 'version.json')) + $L) }
function Detected([string]$pkg) { $d = Detect $pkg; return ($d.Code -eq 0 -and $d.Out -match 'is installed') }
function Short($r) { "exit $($r.Code): " + (($r.Text -split "`r?`n")[0]) }
function PowerPoint([bool]$open) { Write-Json (Join-Path $Base 'processes.json') $(if ($open) { @('POWERPNT') } else { @() }) }

function Task-Xml { try { [string](Export-ScheduledTask -TaskName $TestTask -TaskPath '\' -ErrorAction Stop) } catch { '' } }
function Task-Dll { $t = Get-ScheduledTask -TaskName $TestTask -TaskPath '\' -ErrorAction SilentlyContinue; if ($t) { [string]$t.Actions[0].Arguments } else { '' } }
# Keys with their values, each value as name:kind:data, sorted (a restore may write values in another order).
function Test-Reg { if (Test-Path $TestKey) { @(Get-ChildItem $TestKey -Recurse | Sort-Object Name | ForEach-Object { $k = $_; $k.Name + '=' + ((@($k.GetValueNames() | Sort-Object | ForEach-Object { "${_}:$($k.GetValueKind($_)):$($k.GetValue($_))" })) -join ',') }) -join ';' } else { '' } }
function State { $p = Join-Path $Base 'LocalAppData\PPTK\state.json'; if (Test-Path $p) { [IO.File]::ReadAllText($p) | ConvertFrom-Json } }
function Journal { $p = Join-Path $Base 'LocalAppData\PPTK\transaction.json'; if (Test-Path $p) { [IO.File]::ReadAllText($p) | ConvertFrom-Json } }

# Production state that must never change (read-only).
function Live-Snap {
    $addin = Join-Path $env:APPDATA 'Microsoft\AddIns\PPTK.ppam'
    $addinHash = ''
    if (Test-Path -LiteralPath $addin) { $s = [IO.File]::Open($addin, 'Open', 'Read', 'ReadWrite'); $h = [Security.Cryptography.SHA256]::Create(); $addinHash = [BitConverter]::ToString($h.ComputeHash($s)); $s.Dispose() }
    @(
        'task=' + $(try { [string](Export-ScheduledTask -TaskName 'PPTKKeys Startup' -ErrorAction Stop) } catch { '' })
        'pptkKey=' + (Test-Path 'HKCU:\Software\PPTK')
        'addinReg=' + ((Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK' -ErrorAction SilentlyContinue | Select-Object Path, AutoLoad) | Out-String)
        'addin=' + $addinHash
        'local=' + $(if (Test-Path (Join-Path $env:LOCALAPPDATA 'PPTK')) { @(Get-ChildItem (Join-Path $env:LOCALAPPDATA 'PPTK') -Recurse -Force -File | ForEach-Object { $_.FullName + '|' + $_.Length + '|' + $_.LastWriteTimeUtc.Ticks }) -join ';' } else { '' })
        'helper=' + (@(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'dotnet.exe' -and $_.CommandLine -like '*PPTKKeys.dll*' } | ForEach-Object { $_.ProcessId }) -join ',')
        'otherTestTasks=' + @(Get-ScheduledTask -TaskPath '\' -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like 'PPTK-Test-*' -and $_.TaskName -ne $TestTask }).Count
    ) -join "`n"
}

$liveBefore = Live-Snap
try {
    New-Item -ItemType Directory -Path $Base | Out-Null
    [IO.File]::WriteAllText((Join-Path $Base 'PPTK-TEST-ROOT'), 'isolated PPTK live-adapter test root', $Utf8)
    PowerPoint $false
    $P100 = New-FakePackage 'v100' '1.0.0' 'A' 'A'
    $P101 = New-FakePackage 'v101' '1.0.1' 'B' 'B'
    $P102 = New-FakePackage 'v102' '1.0.2' 'C' 'C'

    # L1 fresh install: real task + HKCU under test names
    $i = Install $P100
    $xml = Task-Xml
    $ok = $i.Code -eq 0 -and (Detected $P100) -and $xml -match [regex]::Escape("versions\1.0.0\PPTKKeys\PPTKKeys.dll") -and $xml -match '<LogonType>InteractiveToken</LogonType>' -and $xml -match '<LogonTrigger>' `
        -and (Get-ItemProperty "$TestKey\Software\PPTK").InstalledVersion -eq '1.0.0' -and (Get-ItemProperty "$TestKey\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK").AutoLoad -eq 1
    Record 'L1' 'fresh install registers the test task and test keys; detect (live) confirms' $ok (Short $i)

    # L2 upgrade with helper change: task updated, helper restart asked
    $i = Install $P101
    $ok = $i.Code -eq 0 -and (Detected $P101) -and (Task-Dll) -match [regex]::Escape('versions\1.0.1\') -and ([IO.File]::ReadAllText((Join-Path $Base 'helper-actions.log')) -match 'stop\r?\nstart')
    Record 'L2' 'upgrade 1.0.0 -> 1.0.1 updates the real task in place' $ok (Short $i)

    # L3 failure after the add-in step: task XML and keys restored exactly
    $xmlBefore = Task-Xml; $regBefore = Test-Reg
    $i = Install $P102 @('-TestFailAt', 'addin')
    Record 'L3' 'failed upgrade: task-xml journal restores the real task and keys exactly' ($i.Code -eq 1 -and (Task-Xml) -eq $xmlBefore -and (Test-Reg) -eq $regBefore -and (Detected $P101) -and -not (Journal)) (Short $i)

    # L4 crash mid-operation: journal is task-xml; the next run undoes it
    $i = Install $P102 @('-TestCrashAt', 'registration')
    $j = Journal
    $crashOk = $i.Code -eq 99 -and $j -and $j.tasksFormat -eq 'task-xml' -and $j.tasksBefore -match 'PPTK-Test-'
    $i2 = Install $P101
    Record 'L4' 'crash leaves a task-xml journal; the next run restores task and keys' ($crashOk -and $i2.Code -eq 0 -and $i2.Out -match 'no change' -and (Task-Xml) -eq $xmlBefore -and (Test-Reg) -eq $regBefore -and (Detected $P101)) ("crash $($i.Code), then " + (Short $i2))

    # L5 PowerPoint open: staged as pending, installed version and task untouched
    PowerPoint $true
    $i = Install $P102
    $st = State
    Record 'L5' 'PowerPoint open: 1618, pending staged, task and keys unchanged' ($i.Code -eq 1618 -and $st.pending.version -eq '1.0.2' -and (Task-Xml) -eq $xmlBefore -and (Test-Reg) -eq $regBefore -and (Detected $P101)) (Short $i)

    # L6 rollback cancels the pending update; then apply; then rollback one step
    $i = Rollback
    $cancel = $i.Code -eq 0 -and -not (State).pending -and (Detected $P101)
    PowerPoint $false
    $i2 = Install $P102
    $applied = $i2.Code -eq 0 -and (Detected $P102)
    $i3 = Rollback
    Record 'L6' 'cancel pending, apply 1.0.2, roll back to 1.0.1 (task follows)' ($cancel -and $applied -and $i3.Code -eq 0 -and (Detected $P101) -and (Task-Dll) -match [regex]::Escape('versions\1.0.1\')) ("$(Short $i) / $(Short $i2) / $(Short $i3)")

    # L7 uninstall failure: everything restored
    $xmlBefore = Task-Xml; $regBefore = Test-Reg
    $i = Uninstall @('-TestFailAt', 'files')
    Record 'L7' 'failed uninstall restores the real task and keys' ($i.Code -eq 1 -and (Task-Xml) -eq $xmlBefore -and (Test-Reg) -eq $regBefore -and (Detected $P101)) (Short $i)

    # L8 uninstall
    $i = Uninstall
    $d = Detect $P101
    Record 'L8' 'uninstall removes the test task and keys; detect says NotInstalled' ($i.Code -eq 0 -and -not (Task-Xml) -and -not (Test-Path "$TestKey\Software\PPTK") -and -not (Test-Path "$TestKey\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK") -and $d.Code -eq 1 -and $d.Err -match 'NotInstalled') ("$(Short $i) / detect $($d.Code)")

    # L13 started by a 32-bit PowerShell (as Intune may do): relaunched as 64-bit, the task runs the 64-bit host
    $script:PsHost = Join-Path $env:windir 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    $i = Install $P100
    $exec = $(try { [string](Get-ScheduledTask -TaskName $TestTask -TaskPath '\' -ErrorAction Stop).Actions[0].Execute } catch { '' })
    $u = Uninstall
    $script:PsHost = 'powershell.exe'
    Record 'L13' '32-bit host: install and uninstall relaunch as 64-bit; task uses the x64 dotnet.exe' ($i.Code -eq 0 -and $exec -and $exec -notmatch '\(x86\)' -and $u.Code -eq 0 -and -not (Task-Xml)) ("$(Short $i) / $exec / $(Short $u)")

    # L9 a journal in the wrong format is refused, never guessed
    New-Item -ItemType Directory -Force -Path (Join-Path $Base 'LocalAppData\PPTK') | Out-Null
    Write-Json (Join-Path $Base 'LocalAppData\PPTK\transaction.json') ([ordered]@{ journalVersion = 1; tasksFormat = 'sim-json'; operation = 'install'; target = '1.0.0'; tasksBefore = '[]'; stateBefore = $null; pptkKeyBefore = $null; addinKeyBefore = $null; addinBefore = $null; createdVersionDir = $null })
    $i = Install $P100
    $kept = [bool](Journal)
    Remove-Item -LiteralPath (Join-Path $Base 'LocalAppData\PPTK\transaction.json') -Force
    Record 'L9' 'sim-json journal in a live environment: RESTORE FAILED, journal kept, nothing installed' ($i.Code -eq 1 -and $i.Err -match 'RESTORE FAILED' -and $kept -and -not (Task-Xml)) (Short $i)

    # L10 production helper discovery (read-only: finds the running helper's message window, sends nothing)
    $probe = @"
`$Root = ''; `$LiveTestId = ''; `$EnvPrefix = 'probe'; function Log(`$m) {}
. '$((Join-Path $RepoRoot 'src\Deploy\deploy-env.ps1').Replace("'", "''"))'
(@(Get-HelperWindows) | ForEach-Object { `$_.Pid } | Sort-Object) -join ','
"@
    $found = powershell -NoProfile -EncodedCommand ([Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($probe)))
    $running = (@(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'dotnet.exe' -and $_.CommandLine -like '*PPTKKeys.dll*' } | ForEach-Object { $_.ProcessId } | Sort-Object) -join ',')
    Record 'L10' 'helper discovery finds exactly the running helper(s) (read-only)' ([string]$found -eq $running) "found '$found', running '$running'"
} finally {
    if (Get-ScheduledTask -TaskName $TestTask -TaskPath '\' -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $TestTask -TaskPath '\' -Confirm:$false }
    if (Test-Path $TestKey) { Remove-Item -Path $TestKey -Recurse -Force }
    if (Test-Path -LiteralPath $Base) { Remove-Item -LiteralPath $Base -Recurse -Force -ErrorAction SilentlyContinue }
}
$liveAfter = Live-Snap
Record 'L11' 'production task, keys, add-in, PPTK folder and helper unchanged' ($liveBefore -eq $liveAfter) $(if ($liveBefore -eq $liveAfter) { 'identical' } else { (Compare-Object ($liveBefore -split "`n") ($liveAfter -split "`n") | Out-String).Trim() })
$left = @()
if (Get-ScheduledTask -TaskName $TestTask -TaskPath '\' -ErrorAction SilentlyContinue) { $left += "task \$TestTask" }
if (Test-Path $TestKey) { $left += $TestKey }
if (Test-Path -LiteralPath $Base) { $left += $Base }
Record 'L12' 'no test leftovers' (-not $left.Count) $(if ($left.Count) { $left -join ', ' } else { 'none' })

$failed = @($results | Where-Object { -not $_.Ok }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
