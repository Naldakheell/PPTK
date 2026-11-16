<#
.SYNOPSIS
  Tests for src\Deploy\install.ps1 (Phase 1B), isolated test roots only.

.DESCRIPTION
  Each case creates a test root (marker PPTK-TEST-ROOT, identity.json, simulated registry, tasks,
  processes and .NET runtime) and runs install.ps1 and the real detect.ps1 as separate processes.
  Fake packages carry small files; case A2 also installs the real package (-Package).
  The live registry, Task Scheduler, profile folders, PPTKKeys and PowerPoint are only read, to
  prove they did not change (X, Y). Exit code 0 when every case passes, 1 otherwise.
#>
param([string]$RepoRoot = '', [string]$Package = '')

$ErrorActionPreference = 'Stop'
if (-not $RepoRoot) { $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path }
$Install = Join-Path $RepoRoot 'src\Deploy\install.ps1'
$Detect = Join-Path $RepoRoot 'src\Deploy\detect.ps1'
$Base = Join-Path ([IO.Path]::GetTempPath()) ('PPTK-install-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $Base | Out-Null
$results = New-Object System.Collections.ArrayList
$Utf8 = New-Object Text.UTF8Encoding($false)
$AddinKey = 'Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK'
$SettingsKey = 'Software\VB and VBA Program Settings\Instrumenta\General'
$LibraryKey = 'Software\VB and VBA Program Settings\PPTK\SlideLibrary'

function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) {
    [void]$results.Add([pscustomobject]@{ Id = $id; Ok = $ok })
    Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail)
}
function Sha([string]$p) { (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLowerInvariant() }
function JsonText($o) { if ($o -is [Array]) { '[' + ((@($o) | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 10 }) -join ',') + ']' } else { ConvertTo-Json -InputObject $o -Depth 10 } }
function Write-Json([string]$p, $o) { [IO.File]::WriteAllText($p, (JsonText $o), $Utf8) }
# ForEach-Object unrolls a JSON array (Windows PowerShell 5.1 emits it as one object).
function Read-Json([string]$p) { if (Test-Path -LiteralPath $p) { [IO.File]::ReadAllText($p) | ConvertFrom-Json | ForEach-Object { $_ } } }

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

function New-Root([string]$name, [switch]$NoRuntime, [string]$RuntimeVersion = '10.0.14', [switch]$MachineRuntime) {
    $r = Join-Path $Base $name
    New-Item -ItemType Directory -Force -Path $r, (Join-Path $r 'Documents\PPTK Flip Backups') | Out-Null
    [IO.File]::WriteAllText((Join-Path $r 'PPTK-TEST-ROOT'), 'isolated PPTK installer test root', $Utf8)
    Write-Json (Join-Path $r 'identity.json') ([ordered]@{ name = 'TESTDOM\tester'; sid = 'S-1-5-21-1000-2000-3000-1001' })
    Write-Json (Join-Path $r 'registry.json') ([ordered]@{ $SettingsKey = [ordered]@{ ContextualButtons = 'True'; OperatingMode = 'pro' }; $LibraryKey = [ordered]@{ Folder = 'D:\Team\Slide Library' } })
    Write-Json (Join-Path $r 'processes.json') @()
    [IO.File]::WriteAllText((Join-Path $r 'Documents\PPTK Flip Backups\deck (before flip).pptx'), 'user backup', $Utf8)
    if (-not $NoRuntime) {
        $dn = if ($MachineRuntime) { Join-Path $r 'ProgramFiles\dotnet' } else { Join-Path $r 'LocalAppData\Microsoft\dotnet' }
        New-Item -ItemType Directory -Force -Path (Join-Path $dn "shared\Microsoft.NETCore.App\$RuntimeVersion") | Out-Null
        [IO.File]::WriteAllText((Join-Path $dn 'dotnet.exe'), 'placeholder for the runtime check; never executed', $Utf8)
    }
    return $r
}

function Invoke-Ps([string]$script, [string[]]$arguments) {
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $script + '" ' + (($arguments | ForEach-Object { if ($_ -match '^-') { $_ } else { '"' + $_ + '"' } }) -join ' ')
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8; $psi.StandardErrorEncoding = [Text.Encoding]::UTF8
    $p = [Diagnostics.Process]::Start($psi); $o = $p.StandardOutput.ReadToEnd(); $e = $p.StandardError.ReadToEnd(); $p.WaitForExit()
    return [pscustomobject]@{ Code = $p.ExitCode; Out = $o.Trim(); Err = $e.Trim(); Text = ($o + ' ' + $e).Trim() }
}
function Install([string]$root, [string]$pkg, [string[]]$extra = @()) { Invoke-Ps $Install (@('-Root', $root, '-PackagePath', $pkg) + $extra) }
function Rollback([string]$root, [string[]]$extra = @()) { Invoke-Ps $Install (@('-Root', $root, '-Rollback') + $extra) }
function Detect([string]$root, [string]$pkg) { Invoke-Ps $Detect @('-Root', $root, '-ManifestPath', (Join-Path $pkg 'version.json')) }
function Detected([string]$root, [string]$pkg) { $d = Detect $root $pkg; return ($d.Code -eq 0 -and $d.Out -match 'is installed') }
function Short($r) { "exit $($r.Code): " + (($r.Text -split "`r?`n")[0]) }

# Everything that matters in a root (files with content hash, plus registry/tasks/state text).
function Snap([string]$root) {
    @(Get-ChildItem -LiteralPath $root -Recurse -File -Force | Where-Object { $_.FullName -notmatch '\\logs\\|state\.lock$' } | ForEach-Object { $_.FullName.Substring($root.Length) + '|' + (Sha $_.FullName) }) -join "`n"
}
function Helper-Log([string]$root) { $p = Join-Path $root 'helper-actions.log'; if (Test-Path -LiteralPath $p) { @([IO.File]::ReadAllLines($p)) } else { @() } }
function Addin([string]$root) { Join-Path $root 'AppData\Microsoft\AddIns\PPTK.ppam' }
function State([string]$root) { Read-Json (Join-Path $root 'LocalAppData\PPTK\state.json') }
function Reg([string]$root) { Read-Json (Join-Path $root 'registry.json') }
function Has-Pending([string]$root) { $s = State $root; return ($null -ne $s -and $s.PSObject.Properties.Name -contains 'pending' -and $null -ne $s.pending) }
function Settings-Intact([string]$root) {
    $r = Reg $root
    return ($r.$SettingsKey.OperatingMode -eq 'pro' -and $r.$SettingsKey.ContextualButtons -eq 'True' -and $r.$LibraryKey.Folder -eq 'D:\Team\Slide Library' -and (Test-Path -LiteralPath (Join-Path $root 'Documents\PPTK Flip Backups\deck (before flip).pptx')))
}

# Live state that must never change (read-only snapshot).
function Live-Snap {
    $addin = Join-Path $env:APPDATA 'Microsoft\AddIns\PPTK.ppam'
    $addinHash = ''
    if (Test-Path -LiteralPath $addin) { $s = [IO.File]::Open($addin, 'Open', 'Read', 'ReadWrite'); $h = [Security.Cryptography.SHA256]::Create(); $addinHash = [BitConverter]::ToString($h.ComputeHash($s)); $s.Dispose() }
    [pscustomobject]@{
        Task = $(try { Export-ScheduledTask -TaskName 'PPTKKeys Startup' -ErrorAction Stop } catch { '' })
        PptkKey = (Test-Path 'HKCU:\Software\PPTK')
        AddinReg = ((Get-ItemProperty -Path ('HKCU:\' + $AddinKey) -ErrorAction SilentlyContinue | Select-Object Path, AutoLoad) | Out-String)
        Addin = $addinHash
        Local = $(if (Test-Path (Join-Path $env:LOCALAPPDATA 'PPTK')) { @(Get-ChildItem (Join-Path $env:LOCALAPPDATA 'PPTK') -Recurse -Force -File | ForEach-Object { $_.FullName + '|' + $_.Length }) -join ';' } else { '' })
        Helper = (@(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'dotnet.exe' -and $_.CommandLine -like '*PPTKKeys.dll*' } | ForEach-Object { $_.ProcessId }) -join ',')
        PowerPoint = (@(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) -join ',')
    }
}

try {
    $liveBefore = Live-Snap

    $P100 = New-FakePackage 'v100' '1.0.0' 'A' 'A'
    $P101 = New-FakePackage 'v101' '1.0.1' 'B' 'A'   # add-in changed only
    $P102 = New-FakePackage 'v102' '1.0.2' 'B' 'B'   # helper changed only
    $P103 = New-FakePackage 'v103' '1.0.3' 'C' 'C'   # both changed
    $P104 = New-FakePackage 'v104' '1.0.4' 'C' 'C'   # neither changed

    # ---------------------------------------------------------------- A, B, W
    $r = New-Root 'A'; $i = Install $r $P100
    $st = State $r
    Record 'A' 'fresh installation' ($i.Code -eq 0 -and (Detected $r $P100) -and $st.installed -eq '1.0.0' -and (Reg $r).'Software\PPTK'.InstalledVersion -eq '1.0.0' -and @(Read-Json (Join-Path $r 'tasks.json')).Count -eq 1) (Short $i)
    if ($Package) {
        $ra = New-Root 'A2'; $i2 = Install $ra $Package
        Record 'A2' 'fresh installation of the real 1.0.0 package' ($i2.Code -eq 0 -and (Detected $ra $Package)) (Short $i2)
    }
    $before = Snap $r; $hl = @(Helper-Log $r).Count; $i = Install $r $P100
    Record 'B' 'repeated installation: success, nothing changed, helper not restarted' ($i.Code -eq 0 -and (Snap $r) -eq $before -and @(Helper-Log $r).Count -eq $hl -and (Detected $r $P100)) (Short $i)

    # ---------------------------------------------------------------- C, D, E, F (upgrades on one root)
    $r = New-Root 'Up'; [void](Install $r $P100)
    $hl = @(Helper-Log $r).Count; $i = Install $r $P101
    Record 'C' 'add-in-only upgrade: add-in replaced, helper not restarted' ($i.Code -eq 0 -and (Detected $r $P101) -and (Sha (Addin $r)) -eq (Sha (Join-Path $P101 'payload\PPTK.ppam')) -and @(Helper-Log $r).Count -eq $hl -and (State $r).previous -eq '1.0.0') (Short $i)
    $addinTime = (Get-Item -LiteralPath (Addin $r)).LastWriteTimeUtc; $hl = @(Helper-Log $r).Count; $i = Install $r $P102
    $acts = @(Helper-Log $r)[$hl..99]
    Record 'D' 'helper-only upgrade: helper restarted, add-in untouched' ($i.Code -eq 0 -and (Detected $r $P102) -and (Get-Item -LiteralPath (Addin $r)).LastWriteTimeUtc -eq $addinTime -and ($acts -join ',') -eq 'stop,start') ("helper actions: " + ($acts -join ','))
    $hl = @(Helper-Log $r).Count; $i = Install $r $P103
    $acts = @(Helper-Log $r)[$hl..99]
    Record 'E' 'both components upgraded' ($i.Code -eq 0 -and (Detected $r $P103) -and (Sha (Addin $r)) -eq (Sha (Join-Path $P103 'payload\PPTK.ppam')) -and ($acts -join ',') -eq 'stop,start') ("helper actions: " + ($acts -join ','))
    $addinTime = (Get-Item -LiteralPath (Addin $r)).LastWriteTimeUtc; $hl = @(Helper-Log $r).Count; $i = Install $r $P104
    $vers = @(Get-ChildItem -LiteralPath (Join-Path $r 'LocalAppData\PPTK\versions') -Directory | ForEach-Object { $_.Name } | Sort-Object)
    Record 'F' 'no-change upgrade: new version recorded, no file replaced, no restart' ($i.Code -eq 0 -and (Detected $r $P104) -and (Get-Item -LiteralPath (Addin $r)).LastWriteTimeUtc -eq $addinTime -and @(Helper-Log $r).Count -eq $hl) ("kept versions: " + ($vers -join ', '))
    Record 'F2' 'old versions pruned only after success; installed + previous kept' (($vers -join ',') -eq '1.0.3,1.0.4') ''

    # ---------------------------------------------------------------- G, H (PowerPoint open)
    $r = New-Root 'Pend'; [void](Install $r $P100)
    Write-Json (Join-Path $r 'processes.json') @('POWERPNT')
    $tasks0 = [IO.File]::ReadAllText((Join-Path $r 'tasks.json'))
    $i = Install $r $P101; $st = State $r
    $pfile = Join-Path $r 'LocalAppData\PPTK\pending\PPTK-1.0.1.ppam'
    Record 'G1' 'PowerPoint open: exit 1618, add-in staged, installed stays 1.0.0, task untouched, 1.0.0 still detected' ($i.Code -eq 1618 -and $st.installed -eq '1.0.0' -and $st.pending.version -eq '1.0.1' -and (Test-Path -LiteralPath $pfile) -and (Sha $pfile) -eq $st.pending.sha256 -and (Reg $r).'Software\PPTK'.InstalledVersion -eq '1.0.0' -and (Sha (Addin $r)) -eq (Sha (Join-Path $P100 'payload\PPTK.ppam')) -and [IO.File]::ReadAllText((Join-Path $r 'tasks.json')) -eq $tasks0 -and (Detected $r $P100)) (Short $i)
    $d = Detect $r $P101
    Record 'G2' 'detect.ps1: 1.0.1 not detected (PendingUpdate)' ($d.Code -ne 0 -and $d.Err -match 'PendingUpdate') (($d.Err -split "`r?`n")[0])
    $sp = Join-Path $r 'LocalAppData\PPTK\state.json'; $o = Read-Json $sp; $o.pending.notified = $true; Write-Json $sp $o
    $i = Install $r $P101; $st = State $r
    Record 'H1' 'rerun while still open: 1618 again, same pending, notified flag kept' ($i.Code -eq 1618 -and $st.pending.version -eq '1.0.1' -and $st.pending.notified -eq $true -and @(Get-ChildItem -LiteralPath (Join-Path $r 'LocalAppData\PPTK\pending') -File).Count -eq 1) (Short $i)
    $lockedAddin = [IO.File]::Open((Addin $r), 'Open', 'Read', 'Read')
    try { Write-Json (Join-Path $r 'processes.json') @(); $i = Install $r $P101 } finally { $lockedAddin.Dispose() }
    Record 'H2' 'add-in file locked (no process listed): still staged, 1618' ($i.Code -eq 1618 -and (State $r).installed -eq '1.0.0') (Short $i)
    $i = Install $r $P101; $st = State $r
    Record 'H3' 'PowerPoint closed: rerun applies the staged version' ($i.Code -eq 0 -and $st.installed -eq '1.0.1' -and $null -eq $st.pending -and (Detected $r $P101) -and -not (Test-Path -LiteralPath $pfile)) (Short $i)

    # ---------------------------------------------------------------- PA..PF (pending: no partial switch)
    $r = New-Root 'PendBoth'; [void](Install $r $P100)
    Write-Json (Join-Path $r 'processes.json') @('POWERPNT')
    $tasks0 = [IO.File]::ReadAllText((Join-Path $r 'tasks.json')); $hl = @(Helper-Log $r).Count
    $i = Install $r $P103; $st = State $r
    Record 'PA' 'both changed, PowerPoint open: task and helper untouched, 1.0.3 staged completely' ($i.Code -eq 1618 -and [IO.File]::ReadAllText((Join-Path $r 'tasks.json')) -eq $tasks0 -and @(Helper-Log $r).Count -eq $hl -and $st.pending.version -eq '1.0.3' -and $st.pending.keysChanged -eq $true -and (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\versions\1.0.3\PPTKKeys\PPTKKeys.dll'))) (Short $i)
    $dOld = Detect $r $P100; $dNew = Detect $r $P103
    Record 'PB' 'during pending: 1.0.0 detected, 1.0.3 not detected (PendingUpdate)' ($dOld.Code -eq 0 -and $dOld.Out -match 'is installed' -and $dNew.Code -ne 0 -and $dNew.Err -match 'PendingUpdate') ("old exit $($dOld.Code), new: " + (($dNew.Err -split "`r?`n")[0]))
    Write-Json (Join-Path $r 'processes.json') @(); $clean = Snap $r
    $x = Install $r $P103 @('-TestFailAt', 'addin')
    Record 'PC' 'apply fails mid-way (after task + add-in): all undone, 1.0.0 detected, update still pending' ($x.Code -eq 1 -and (Detected $r $P100) -and -not (Detected $r $P103) -and (State $r).pending.version -eq '1.0.3' -and [IO.File]::ReadAllText((Join-Path $r 'tasks.json')) -eq $tasks0) (Short $x)
    $hl = @(Helper-Log $r).Count; $x = Install $r $P103; $t = @(Read-Json (Join-Path $r 'tasks.json'))[0]; $acts = @(Helper-Log $r)[$hl..99]
    Record 'PD' 'PowerPoint closed: task, helper, add-in and state switch together; 1.0.0 no longer detected' ($x.Code -eq 0 -and (Detected $r $P103) -and -not (Detected $r $P100) -and $t.actions[0].arguments -match 'versions\\1\.0\.3\\PPTKKeys' -and ($acts -join ',') -eq 'stop,start' -and $null -eq (State $r).pending -and (Sha (Addin $r)) -eq (Sha (Join-Path $P103 'payload\PPTK.ppam'))) ("helper actions: " + ($acts -join ','))
    $r = New-Root 'PendKeys'; [void](Install $r $P101)
    Write-Json (Join-Path $r 'processes.json') @('POWERPNT'); $hl = @(Helper-Log $r).Count
    $x = Install $r $P102; $acts = @(Helper-Log $r)[$hl..99]
    Record 'PE' 'helper-only update with PowerPoint open: applied now (exit 0), helper restarted' ($x.Code -eq 0 -and (Detected $r $P102) -and ($acts -join ',') -eq 'stop,start' -and $null -eq (State $r).pending) ("helper actions: " + ($acts -join ','))
    $r = New-Root 'PendCrash'; [void](Install $r $P100)
    Write-Json (Join-Path $r 'processes.json') @('POWERPNT')
    $x = Install $r $P103 @('-TestFailAt', 'state')
    Record 'PF' 'failure while staging a pending update: state restored, 1.0.0 detected, nothing pending' ($x.Code -eq 1 -and (Detected $r $P100) -and -not (Has-Pending $r)) (Short $x)

    # ---------------------------------------------------------------- RP (rollback with a pending update)
    $r = New-Root 'RollPend'; [void](Install $r $P100); [void](Install $r $P101)
    Write-Json (Join-Path $r 'processes.json') @('POWERPNT'); [void](Install $r $P103)
    $tasks0 = [IO.File]::ReadAllText((Join-Path $r 'tasks.json')); $addin0 = Sha (Addin $r); $hl = @(Helper-Log $r).Count
    $x = Rollback $r; $st = State $r
    Record 'RP1' 'rollback with 1.0.3 pending: pending cancelled, 1.0.1 kept and detected (not 1.0.0)' ($x.Code -eq 0 -and $x.Out -match 'cancelled' -and $st.installed -eq '1.0.1' -and $st.previous -eq '1.0.0' -and $null -eq $st.pending -and (Detected $r $P101) -and -not (Detected $r $P100) -and -not (Detected $r $P103)) (Short $x)
    Record 'RP2' 'cancel touched nothing else: task, add-in, helper unchanged; pending files and versions\1.0.3 removed' ([IO.File]::ReadAllText((Join-Path $r 'tasks.json')) -eq $tasks0 -and (Sha (Addin $r)) -eq $addin0 -and @(Helper-Log $r).Count -eq $hl -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\pending')) -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\versions\1.0.3')) -and (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\versions\1.0.0'))) ''
    Write-Json (Join-Path $r 'processes.json') @()
    $x = Rollback $r
    Record 'RP3' 'after the cancel, a normal rollback still goes one step back to 1.0.0' ($x.Code -eq 0 -and (Detected $r $P100) -and (State $r).rolledBackFrom -eq '1.0.1') (Short $x)

    # ---------------------------------------------------------------- RH (rollback history)
    $r = New-Root 'RollHist'; [void](Install $r $P100); [void](Install $r $P103)
    $x = Rollback $r; $st = State $r
    Record 'RH1' 'rollback records rolledBackFrom 1.0.3 and clears previous' ($x.Code -eq 0 -and $st.installed -eq '1.0.0' -and $st.rolledBackFrom -eq '1.0.3' -and $null -eq $st.previous -and (Detected $r $P100)) (Short $x)
    $clean = Snap $r; $x = Rollback $r
    Record 'RH2' 'second rollback refused (no ping-pong to 1.0.3), nothing changed, 1.0.0 detected' ($x.Code -eq 1 -and $x.Err -match 'RollbackUnavailable.*already rolled back from 1\.0\.3' -and (Snap $r) -eq $clean -and (Detected $r $P100) -and -not (Detected $r $P103)) (Short $x)
    $x = Install $r $P103; $st = State $r
    Record 'RH3' 'reinstalling 1.0.3 after a rollback works; rolledBackFrom cleared, previous 1.0.0' ($x.Code -eq 0 -and (Detected $r $P103) -and $st.previous -eq '1.0.0' -and -not $st.rolledBackFrom) (Short $x)
    $sp = Join-Path $r 'LocalAppData\PPTK\state.json'; $o = Read-Json $sp; $o.previous = '1.0.3'; Write-Json $sp $o
    $clean = Snap $r; $x = Rollback $r
    Record 'RH4' 'invalid previous (not older than installed): refused, nothing changed' ($x.Code -eq 1 -and $x.Err -match 'RollbackUnavailable.*not older' -and (Snap $r) -eq $clean -and (Detected $r $P103)) (Short $x)

    # ---------------------------------------------------------------- I, J, K, L, M, Z (invalid input: nothing changes)
    $r = New-Root 'Bad'; [void](Install $r $P100); $clean = Snap $r
    $bad = { param($id, $name, $pkgDir, $pattern) $x = Install $r $pkgDir; Record $id $name ($x.Code -eq 1 -and $x.Err -match $pattern -and (Snap $r) -eq $clean -and (Detected $r $P100)) (Short $x) }
    & $bad 'I1' 'manifest missing' (Join-Path $Base 'no-such-package') 'InvalidPackage: version.json not found'
    $q = New-FakePackage 'malformed' '1.0.5' 'D' 'D'; [IO.File]::WriteAllText((Join-Path $q 'version.json'), '{ "packageName": ', $Utf8)
    & $bad 'I2' 'manifest malformed' $q 'not valid JSON'
    $q = New-FakePackage 'badver' '1.0.5' 'D' 'D'; $m = Read-Json (Join-Path $q 'version.json'); $m.packageVersion = 'latest'; Write-Json (Join-Path $q 'version.json') $m
    & $bad 'I3' 'manifest version invalid' $q 'packageVersion'
    $q = New-FakePackage 'missingfile' '1.0.5' 'D' 'D'; Remove-Item -LiteralPath (Join-Path $q 'payload\PPTKKeys\PPTKKeys.deps.json')
    & $bad 'J' 'payload file missing' $q 'payload file missing: payload/PPTKKeys/PPTKKeys.deps.json'
    $q = New-FakePackage 'corrupt' '1.0.5' 'D' 'D'; [IO.File]::AppendAllText((Join-Path $q 'payload\PPTKKeys\PPTKKeys.dll'), 'x')
    & $bad 'K' 'corrupt SHA-256' $q 'SHA-256 mismatch: payload/PPTKKeys/PPTKKeys.dll'
    $q = New-FakePackage 'traversal' '1.0.5' 'D' 'D'; $m = Read-Json (Join-Path $q 'version.json'); $m.files[0].path = '../../AppData/Microsoft/AddIns/PPTK.ppam'; Write-Json (Join-Path $q 'version.json') $m
    & $bad 'L' 'path traversal in the manifest' $q 'unsafe or missing file path'
    $q = New-FakePackage 'withexe' '1.0.5' 'D' 'D'; [IO.File]::WriteAllText((Join-Path $q 'payload\PPTKKeys\PPTKKeys.exe'), 'not a program', $Utf8)
    & $bad 'Z1' 'unlisted EXE in the package' $q 'not listed in version\.json: payload/PPTKKeys/PPTKKeys\.exe'
    $q = New-FakePackage 'listedexe' '1.0.5' 'D' 'D'; $p = Join-Path $q 'payload\tool.exe'; [IO.File]::WriteAllText($p, 'x', $Utf8); $m = Read-Json (Join-Path $q 'version.json'); $m.files += [pscustomobject]@{ path = 'payload/tool.exe'; size = 1; sha256 = (Sha $p) }; Write-Json (Join-Path $q 'version.json') $m
    & $bad 'Z2' 'EXE listed in the manifest' $q 'executable or banned file type'
    $d0 = New-FakePackage 'older' '0.9.0' 'Z' 'Z'
    & $bad 'I4' 'downgrade refused (rollback is the way back)' $d0 'Downgrade'

    $r = New-Root 'NoNet' -NoRuntime; $x = Install $r $P100
    Record 'M1' '.NET runtime missing: refused, nothing installed' ($x.Code -eq 1 -and $x.Err -match 'MissingRuntime' -and -not (Test-Path -LiteralPath (Addin $r)) -and -not (Test-Path -LiteralPath (Join-Path $r 'tasks.json'))) (Short $x)
    $r = New-Root 'Net9' -RuntimeVersion '9.0.5'; $x = Install $r $P100
    Record 'M2' '.NET 9 only (wrong major): refused' ($x.Code -eq 1 -and $x.Err -match 'MissingRuntime') (Short $x)
    $r = New-Root 'NetMachine' -MachineRuntime; $x = Install $r $P100
    $t = @(Read-Json (Join-Path $r 'tasks.json'))[0]
    Record 'M3' 'machine-wide .NET (Program Files) is found and used' ($x.Code -eq 0 -and (Detected $r $P100) -and $t.actions[0].execute -like '*ProgramFiles\dotnet\dotnet.exe') (Short $x)

    # ---------------------------------------------------------------- N (missing task repaired)
    $r = New-Root 'NoTask'; [void](Install $r $P100); Write-Json (Join-Path $r 'tasks.json') @()
    $x = Install $r $P100
    Record 'N' 'missing scheduled task: re-created by a repair run' ($x.Code -eq 0 -and (Detected $r $P100) -and @(Read-Json (Join-Path $r 'tasks.json')).Count -eq 1) (Short $x)

    # ---------------------------------------------------------------- O, P (legacy and unknown states)
    $mkLegacy = { param($name, [string]$taskUser = 'tester')
        $lr = New-Root $name
        $ad = Join-Path $lr 'AppData\Microsoft\AddIns'; New-Item -ItemType Directory -Force -Path $ad | Out-Null
        [IO.File]::WriteAllText((Join-Path $ad 'PPTK.ppam'), 'manually installed add-in', $Utf8)
        $lk = Join-Path $lr 'LocalAppData\PPTK\PPTKKeys'; New-Item -ItemType Directory -Force -Path $lk | Out-Null
        [IO.File]::WriteAllText((Join-Path $lk 'PPTKKeys.dll'), 'manual helper', $Utf8)
        $reg = Read-Json (Join-Path $lr 'registry.json'); $reg | Add-Member -NotePropertyName $AddinKey -NotePropertyValue ([ordered]@{ Path = 'PPTK.ppam'; AutoLoad = 1 }); Write-Json (Join-Path $lr 'registry.json') $reg
        Write-Json (Join-Path $lr 'tasks.json') @([ordered]@{ name = 'PPTKKeys Startup'; path = '\'; enabled = $true; userId = $taskUser; logonType = 'Interactive'; runLevel = 'Limited'
            triggers = @([ordered]@{ type = 'Logon'; userId = 'TESTDOM\tester'; enabled = $true })
            actions = @([ordered]@{ execute = (Join-Path $lr 'LocalAppData\Microsoft\dotnet\dotnet.exe'); arguments = 'exec "' + (Join-Path $lk 'PPTKKeys.dll') + '"' }) })
        $lr }
    $r = & $mkLegacy 'Legacy'
    $d = Detect $r $P100
    $x = Install $r $P100; $st = State $r; $tasks = @(Read-Json (Join-Path $r 'tasks.json'))
    $lb = if ($st.legacyBackup) { Join-Path $r ('LocalAppData\PPTK\' + $st.legacyBackup) } else { '' }
    Record 'O1' 'manual installation adopted: one task, detected, legacy add-in and task backed up' ($d.Err -match 'NotInstalled' -and $x.Code -eq 0 -and (Detected $r $P100) -and $tasks.Count -eq 1 -and $lb -and ([IO.File]::ReadAllText((Join-Path $lb 'PPTK.ppam'))) -eq 'manually installed add-in' -and (Test-Path -LiteralPath (Join-Path $lb 'task.json')) -and (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\PPTKKeys\PPTKKeys.dll')) -and (Settings-Intact $r)) (Short $x)
    $r = & $mkLegacy 'LegacyOther' 'OTHERDOM\someone'; $clean = Snap $r; $x = Install $r $P100
    Record 'O2' 'unknown state (task of another user): manual approval required, nothing changed' ($x.Code -eq 1 -and $x.Err -match 'ManualApprovalRequired' -and (Snap $r) -eq $clean) (Short $x)
    $r = & $mkLegacy 'LegacyReg'; $reg = Read-Json (Join-Path $r 'registry.json'); $reg.$AddinKey.Path = 'C:\Other\PPTK.ppam'; Write-Json (Join-Path $r 'registry.json') $reg; $clean = Snap $r; $x = Install $r $P100
    Record 'O3' 'unknown state (registration elsewhere): manual approval required, nothing changed' ($x.Code -eq 1 -and $x.Err -match 'ManualApprovalRequired' -and (Snap $r) -eq $clean) (Short $x)
    $r = New-Root 'Dup'; [void](Install $r $P100)
    $t = @(Read-Json (Join-Path $r 'tasks.json')); $dup = $t[0].PSObject.Copy(); $dup.name = 'PPTKKeys Old'; Write-Json (Join-Path $r 'tasks.json') @($t[0], $dup)
    $clean = Snap $r; $x = Install $r $P101
    Record 'P' 'duplicate PPTKKeys task: refused (manual approval), no second task created' ($x.Code -eq 1 -and $x.Err -match 'other tasks already run PPTKKeys' -and (Snap $r) -eq $clean) (Short $x)

    # ---------------------------------------------------------------- Q, R (rollback)
    $r = New-Root 'Fail'; [void](Install $r $P100); $clean = Snap $r
    $x = Install $r $P103 @('-TestFailAt', 'addin')
    Record 'Q1' 'failure during upgrade: previous version restored and detected, new not detected' ($x.Code -eq 1 -and (Detected $r $P100) -and -not (Detected $r $P103) -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\versions\1.0.3')) -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\transaction.json'))) (Short $x)
    $r = New-Root 'Roll'; [void](Install $r $P100); [void](Install $r $P103)
    $x = Rollback $r; $t = @(Read-Json (Join-Path $r 'tasks.json'))[0]
    Record 'Q2' '-Rollback: previous version restored and detected with its own manifest' ($x.Code -eq 0 -and (Detected $r $P100) -and (State $r).installed -eq '1.0.0' -and (Sha (Addin $r)) -eq (Sha (Join-Path $P100 'payload\PPTK.ppam')) -and $t.actions[0].arguments -match 'versions\\1\.0\.0\\PPTKKeys' -and (Settings-Intact $r)) (Short $x)
    $r = New-Root 'RollNone'; [void](Install $r $P100); $x = Rollback $r
    Record 'Q3' 'rollback without a previous version: refused, install still healthy' ($x.Code -eq 1 -and $x.Err -match 'RollbackUnavailable' -and (Detected $r $P100)) (Short $x)
    $r = New-Root 'RollGone'; [void](Install $r $P100); [void](Install $r $P103); Remove-Item -LiteralPath (Join-Path $r 'LocalAppData\PPTK\versions\1.0.0') -Recurse -Force; $x = Rollback $r
    Record 'Q4' 'rollback version folder missing: refused, current still detected' ($x.Code -eq 1 -and $x.Err -match 'RollbackUnavailable' -and (Detected $r $P103)) (Short $x)
    $r = New-Root 'RollBad'; [void](Install $r $P100); [void](Install $r $P103); [IO.File]::AppendAllText((Join-Path $r 'LocalAppData\PPTK\versions\1.0.0\PPTK.ppam'), 'x'); $clean = Snap $r; $x = Rollback $r
    Record 'R' 'rollback with a corrupt backup: refused, nothing changed, current still detected' ($x.Code -eq 1 -and $x.Err -match 'RollbackCorrupt' -and (Snap $r) -eq $clean -and (Detected $r $P103)) (Short $x)

    # ---------------------------------------------------------------- S (interrupted operations)
    $r = New-Root 'Crash'; [void](Install $r $P100)
    $x = Install $r $P103 @('-TestCrashAt', 'task')
    $mid = Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\transaction.json')
    $y = Install $r $P103
    Record 'S1' 'install killed after the task step: next run undoes it and completes' ($x.Code -eq 99 -and $mid -and $y.Code -eq 0 -and (Detected $r $P103) -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\transaction.json'))) ("crash exit $($x.Code), journal left: $mid; " + (Short $y))
    $r = New-Root 'CrashRoll'; [void](Install $r $P100); [void](Install $r $P103)
    $x = Rollback $r @('-TestCrashAt', 'addin'); $y = Rollback $r
    Record 'S2' 'rollback killed after the add-in step: next rollback completes' ($x.Code -eq 99 -and $y.Code -eq 0 -and (Detected $r $P100)) (Short $y)
    $r = New-Root 'CrashState'; [void](Install $r $P100)
    $x = Install $r $P101 @('-TestCrashAt', 'state'); $y = Install $r $P100
    Record 'S3' 'killed after writing state: next run restores 1.0.0 first (journal), then reinstalls it' ($x.Code -eq 99 -and $y.Code -eq 0 -and (Detected $r $P100) -and (State $r).installed -eq '1.0.0') (Short $y)

    # ---------------------------------------------------------------- T (concurrency)
    $r = New-Root 'Lock'; [void](Install $r $P100); $clean = Snap $r
    $held = [IO.File]::Open((Join-Path $r 'LocalAppData\PPTK\state.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    try { $x = Install $r $P103 @('-LockTimeoutSeconds', '2') } finally { $held.Dispose() }
    Record 'T1' 'second operation while state.lock is held: 1618 Busy, nothing changed' ($x.Code -eq 1618 -and $x.Err -match 'Busy' -and (Snap $r) -eq $clean) (Short $x)
    $x = Install $r $P103
    Record 'T2' 'lock released: the same install now succeeds (no stale marker)' ($x.Code -eq 0 -and (Detected $r $P103)) (Short $x)

    # ---------------------------------------------------------------- MV (transient folder-move failures), TX (leftover txn\)
    $logText = { param($root) (Get-ChildItem -LiteralPath (Join-Path $root 'LocalAppData\PPTK\logs') -File | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n" }
    $r = New-Root 'MoveRetry'; $x = Install $r $P100 @('-TestMoveFailures', '3')
    Record 'MV1' 'staging folder move fails 3 times (transient): retried, installed' ($x.Code -eq 0 -and (Detected $r $P100) -and (& $logText $r) -match 'attempt 3 failed') (Short $x)
    $r = New-Root 'MoveFail'; $x = Install $r $P100 @('-TestMoveFailures', '5')
    Record 'MV2' 'staging folder move keeps failing (5): exit 1, nothing installed, no journal' ($x.Code -eq 1 -and $x.Err -match 'simulated transient failure' -and -not (Test-Path -LiteralPath (Addin $r)) -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\versions\1.0.0')) -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\state.json')) -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\transaction.json'))) (Short $x)
    $r = New-Root 'Leftover'; [void](Install $r $P100); New-Item -ItemType Directory -Force -Path (Join-Path $r 'LocalAppData\PPTK\txn\removed') | Out-Null
    $x = Install $r $P101
    Record 'TX' 'leftover txn\ without a journal (committed run): removed, never undone' ($x.Code -eq 0 -and (Detected $r $P101) -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\txn'))) (Short $x)

    # ---------------------------------------------------------------- RR (file replace retry)
    $r = New-Root 'RepRetry'; [void](Install $r $P100); $x = Install $r $P101 @('-TestReplaceFailures', '4')
    Record 'RR1' 'file replacement fails 4 times (transient 0x80070497): retried, installed' ($x.Code -eq 0 -and (Detected $r $P101) -and (& $logText $r) -match 'attempt 4 failed \(0x80070497') (Short $x)
    $r = New-Root 'RepAddin'; [void](Install $r $P100); $x = Install $r $P101 @('-TestReplaceFailures', '2', '-TestReplaceTarget', 'PPTK.ppam')
    Record 'RR2' 'add-in replacement fails twice (transient): retried, installed' ($x.Code -eq 0 -and (Detected $r $P101) -and (& $logText $r) -match 'replace PPTK\.ppam: attempt 2 failed') (Short $x)
    $r = New-Root 'RepExhaust'; [void](Install $r $P100); $x = Install $r $P101 @('-TestReplaceFailures', '5')
    Record 'RR3' 'file replacement keeps failing (5): exit 1, 1.0.0 restored and detected' ($x.Code -eq 1 -and (Detected $r $P100) -and (& $logText $r) -match 'attempt 5 of 5 failed; giving up' -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK\transaction.json'))) (Short $x)
    $r = New-Root 'RepReadOnly'; [void](Install $r $P100); $reg = Get-Item -LiteralPath (Join-Path $r 'registry.json')
    $reg.IsReadOnly = $true
    try { $x = Install $r $P101 } finally { $reg.IsReadOnly = $false }
    Record 'RR4' 'read-only target (permanent access denied): failed at once, no retry, clear error' ($x.Code -eq 1 -and (& $logText $r) -notmatch 'replace registry\.json: attempt \d+ failed') (Short $x)
    $y = Install $r $P101
    Record 'RR5' 'permission fixed: the next run recovers from the kept journal and installs' ($y.Code -eq 0 -and (Detected $r $P101)) (Short $y)

    # ---------------------------------------------------------------- U, V
    Record 'U' 'user settings, SlideLibrary folder and Documents backups preserved everywhere' ((Settings-Intact (Join-Path $Base 'Up')) -and (Settings-Intact (Join-Path $Base 'Roll')) -and (Settings-Intact (Join-Path $Base 'Legacy')) -and (Settings-Intact (Join-Path $Base 'Pend'))) ''
    $arabic = -join (@(0x0627, 0x062E, 0x062A, 0x0628, 0x0627, 0x0631, 0x0020, 0x2713) | ForEach-Object { [char]$_ })
    $r = New-Root $arabic; $x = Install $r $P100; $y = Install $r $P103; $z = Rollback $r
    Record 'V' 'Arabic/Unicode root: install, upgrade, rollback' ($x.Code -eq 0 -and $y.Code -eq 0 -and $z.Code -eq 0 -and (Detected $r $P100)) ("$($x.Code)/$($y.Code)/$($z.Code)")

    # ---------------------------------------------------------------- X, Y (nothing live touched)
    $code = @([IO.File]::ReadAllLines($Install) | Where-Object { $_ -notmatch '^\s*#' })
    $forbidden = @($code | Where-Object { $_ -match '(?i)Register-ScheduledTask|Set-ScheduledTask|Start-ScheduledTask|Stop-ScheduledTask|Unregister-ScheduledTask|Get-ScheduledTask|Stop-Process|Start-Process|Get-Process|Set-ItemProperty|New-ItemProperty|Remove-ItemProperty|HKCU:|HKEY_CURRENT_USER|Registry::|\$env:(APPDATA|LOCALAPPDATA|USERPROFILE)|GetFolderPath' })
    Record 'Y1' 'install.ps1 has no live registry, task, process or profile operations' ($forbidden.Count -eq 0) ($forbidden -join ' / ')
    $liveAfter = Live-Snap
    $diff = @('Task', 'PptkKey', 'AddinReg', 'Addin', 'Local') | Where-Object { $liveBefore.$_ -ne $liveAfter.$_ }
    Record 'X' 'live task, HKCU, AddIns file and LocalAppData\PPTK unchanged' ($diff.Count -eq 0) ("changed: " + ($diff -join ', '))
    Record 'Y2' 'running PPTKKeys and PowerPoint processes unchanged' ($liveBefore.Helper -eq $liveAfter.Helper -and $liveBefore.PowerPoint -eq $liveAfter.PowerPoint) ("helper pid " + $liveAfter.Helper)
} catch {
    Record 'ERR' 'test run' $false ($_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    Remove-Item -LiteralPath $Base -Recurse -Force -ErrorAction SilentlyContinue
}

$failed = @($results | Where-Object { -not $_.Ok })
Write-Host ("TOTAL {0} tests, {1} failed" -f $results.Count, $failed.Count)
if ($failed.Count) { exit 1 } else { exit 0 }
