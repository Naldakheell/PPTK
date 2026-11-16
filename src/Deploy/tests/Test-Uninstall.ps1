<#
.SYNOPSIS
  Tests for src\Deploy\uninstall.ps1 (Phase 1C), isolated test roots only.

.DESCRIPTION
  Each case builds a test root, installs fake packages with the real install.ps1, then runs
  uninstall.ps1 and the real detect.ps1 as separate processes. The live registry, Task Scheduler,
  profile folders, PPTKKeys and PowerPoint are only read, to prove they did not change (V).
  Exit code 0 when every case passes, 1 otherwise.
#>
param([string]$RepoRoot = '')

$ErrorActionPreference = 'Stop'
if (-not $RepoRoot) { $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path }
$Install = Join-Path $RepoRoot 'src\Deploy\install.ps1'
$Uninstall = Join-Path $RepoRoot 'src\Deploy\uninstall.ps1'
$Detect = Join-Path $RepoRoot 'src\Deploy\detect.ps1'
$Base = Join-Path ([IO.Path]::GetTempPath()) ('PPTK-uninstall-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $Base | Out-Null
$results = New-Object System.Collections.ArrayList
$Utf8 = New-Object Text.UTF8Encoding($false)
$AddinKey = 'Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK'
$OtherAddinKey = 'Software\Microsoft\Office\16.0\PowerPoint\AddIns\OtherTool'
$SettingsKey = 'Software\VB and VBA Program Settings\Instrumenta\General'
$LibraryKey = 'Software\VB and VBA Program Settings\PPTK\SlideLibrary'

function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) {
    [void]$results.Add([pscustomobject]@{ Id = $id; Ok = $ok })
    Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail)
}
function Sha([string]$p) { (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLowerInvariant() }
function JsonText($o) { if ($o -is [Array]) { '[' + ((@($o) | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 10 }) -join ',') + ']' } else { ConvertTo-Json -InputObject $o -Depth 10 } }
function Write-Json([string]$p, $o) { [IO.File]::WriteAllText($p, (JsonText $o), $Utf8) }
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

# Test root with user data that must survive: VBA settings, SlideLibrary folder + registry value,
# Documents backups and decks, an unrelated add-in (file + registration) and an unrelated task.
function New-Root([string]$name) {
    $r = Join-Path $Base $name
    New-Item -ItemType Directory -Force -Path $r, (Join-Path $r 'Documents\PPTK Flip Backups'), (Join-Path $r 'Team\Slide Library'), (Join-Path $r 'AppData\Microsoft\AddIns') | Out-Null
    [IO.File]::WriteAllText((Join-Path $r 'PPTK-TEST-ROOT'), 'isolated PPTK uninstaller test root', $Utf8)
    Write-Json (Join-Path $r 'identity.json') ([ordered]@{ name = 'TESTDOM\tester'; sid = 'S-1-5-21-1000-2000-3000-1001' })
    Write-Json (Join-Path $r 'registry.json') ([ordered]@{
        $SettingsKey = [ordered]@{ ContextualButtons = 'True'; OperatingMode = 'pro' }
        $LibraryKey = [ordered]@{ Folder = (Join-Path $r 'Team\Slide Library') }
        $OtherAddinKey = [ordered]@{ Path = 'OtherTool.ppam'; AutoLoad = 1 } })
    Write-Json (Join-Path $r 'processes.json') @()
    Write-Json (Join-Path $r 'tasks.json') @([ordered]@{ name = 'Backup Sync'; path = '\'; enabled = $true; userId = 'TESTDOM\tester'; logonType = 'Interactive'; runLevel = 'Limited'; triggers = @(); actions = @([ordered]@{ execute = 'C:\Tools\sync.cmd'; arguments = '' }) })
    [IO.File]::WriteAllText((Join-Path $r 'Documents\PPTK Flip Backups\deck (before flip).pptx'), 'user backup', $Utf8)
    [IO.File]::WriteAllText((Join-Path $r 'Documents\Quarterly review.pptx'), 'user deck', $Utf8)
    [IO.File]::WriteAllText((Join-Path $r 'Team\Slide Library\Agenda.pptx'), 'library slide', $Utf8)
    [IO.File]::WriteAllText((Join-Path $r 'AppData\Microsoft\AddIns\OtherTool.ppam'), 'another add-in', $Utf8)
    $dn = Join-Path $r 'LocalAppData\Microsoft\dotnet'
    New-Item -ItemType Directory -Force -Path (Join-Path $dn 'shared\Microsoft.NETCore.App\10.0.14') | Out-Null
    [IO.File]::WriteAllText((Join-Path $dn 'dotnet.exe'), 'placeholder for the runtime check; never executed', $Utf8)
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
function Uninstall([string]$root, [string[]]$extra = @()) { Invoke-Ps $Uninstall (@('-Root', $root) + $extra) }
function Detect([string]$root, [string]$pkg) { Invoke-Ps $Detect @('-Root', $root, '-ManifestPath', (Join-Path $pkg 'version.json')) }
function Detected([string]$root, [string]$pkg) { $d = Detect $root $pkg; return ($d.Code -eq 0 -and $d.Out -match 'is installed') }
# detect.ps1 reports NotInstalled and no leftover piece of PPTK.
function Gone([string]$root, [string]$pkg) { $d = Detect $root $pkg; return ($d.Code -ne 0 -and $d.Err -match 'NotInstalled' -and $d.Err -notmatch 'found without installer markers') }
function Short($r) { "exit $($r.Code): " + (($r.Text -split "`r?`n")[0]) }

# Content of a root, minus logs, the lock file and the simulated helper log (which records actions).
# -Restored: after an undo, state.json carries lastError and registry.json may list restored keys in
# another order; registry.json is then compared by its values instead of its text.
function Snap([string]$root, [switch]$Restored) {
    $files = @(Get-ChildItem -LiteralPath $root -Recurse -File -Force | Where-Object { $_.FullName -notmatch '\\logs\\|state\.lock$|helper-actions\.log$' -and -not ($Restored -and $_.Name -in 'state.json', 'registry.json') } | ForEach-Object { $_.FullName.Substring($root.Length) + '|' + (Sha $_.FullName) }) -join "`n"
    if (-not $Restored) { return $files }
    $reg = Reg $root
    return $files + "`n" + ((@($reg.PSObject.Properties | Sort-Object Name | ForEach-Object { $_.Name + '=' + (ConvertTo-Json -InputObject $_.Value -Compress) })) -join "`n")
}
function State-Json([string]$root) { Read-Json (Pptk $root 'state.json') }
function Helper-Log([string]$root) { $p = Join-Path $root 'helper-actions.log'; if (Test-Path -LiteralPath $p) { @([IO.File]::ReadAllLines($p)) } else { @() } }
function Addin([string]$root) { Join-Path $root 'AppData\Microsoft\AddIns\PPTK.ppam' }
function Pptk([string]$root, [string]$rel) { Join-Path $root "LocalAppData\PPTK\$rel" }
function Reg([string]$root) { Read-Json (Join-Path $root 'registry.json') }
function Tasks([string]$root) { @(Read-Json (Join-Path $root 'tasks.json') | Where-Object { $_ }) }
# Every PPTK-managed piece is gone; the unrelated task stays.
function Removed([string]$root) {
    $reg = Reg $root; $t = Tasks $root
    return (-not (Test-Path -LiteralPath (Addin $root)) -and -not ($reg.PSObject.Properties.Name -contains $AddinKey) -and -not ($reg.PSObject.Properties.Name -contains 'Software\PPTK') -and
        -not (Test-Path -LiteralPath (Pptk $root 'versions')) -and -not (Test-Path -LiteralPath (Pptk $root 'pending')) -and -not (Test-Path -LiteralPath (Pptk $root 'state.json')) -and
        -not (Test-Path -LiteralPath (Pptk $root 'txn')) -and -not (Test-Path -LiteralPath (Pptk $root 'transaction.json')) -and
        @($t | Where-Object { $_.name -eq 'PPTKKeys Startup' }).Count -eq 0)
}
function UserData-Intact([string]$root) {
    $r = Reg $root
    return ($r.$SettingsKey.OperatingMode -eq 'pro' -and $r.$SettingsKey.ContextualButtons -eq 'True' -and $r.$LibraryKey.Folder -eq (Join-Path $root 'Team\Slide Library') -and
        (Test-Path -LiteralPath (Join-Path $root 'Documents\PPTK Flip Backups\deck (before flip).pptx')) -and (Test-Path -LiteralPath (Join-Path $root 'Documents\Quarterly review.pptx')) -and
        [IO.File]::ReadAllText((Join-Path $root 'Team\Slide Library\Agenda.pptx')) -eq 'library slide')
}
function Unrelated-Intact([string]$root) {
    $r = Reg $root
    return ([IO.File]::ReadAllText((Join-Path $root 'AppData\Microsoft\AddIns\OtherTool.ppam')) -eq 'another add-in' -and $r.$OtherAddinKey.Path -eq 'OtherTool.ppam' -and @(Tasks $root | Where-Object { $_.name -eq 'Backup Sync' }).Count -eq 1)
}
function Set-Processes([string]$root, [string[]]$names) { Write-Json (Join-Path $root 'processes.json') @($names) }

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
    $P101 = New-FakePackage 'v101' '1.0.1' 'B' 'A'
    $P103 = New-FakePackage 'v103' '1.0.3' 'C' 'C'

    # ---------------------------------------------------------------- safety boundary
    $x = Invoke-Ps $Uninstall @('-LiveTestId', '0123456789ab')   # never without -Root: that is the production mode
    $nm = Join-Path $Base 'NoMarker'; New-Item -ItemType Directory -Path $nm | Out-Null; [IO.File]::WriteAllText((Join-Path $nm 'identity.json'), '{"name":"TESTDOM\\tester","sid":"S-1"}', $Utf8)
    $y = Uninstall $nm
    Record 'Z' '-LiveTestId without -Root / root without PPTK-TEST-ROOT: exit 2, nothing changed' ($x.Code -eq 2 -and $y.Code -eq 2 -and @(Get-ChildItem -LiteralPath $nm -Recurse -Force).Count -eq 1) ("$($x.Code)/$($y.Code)")

    # ---------------------------------------------------------------- A, U, B, C
    $r = New-Root 'A'; [void](Install $r $P100); $ok0 = Detected $r $P100; $hl = @(Helper-Log $r).Count
    $x = Uninstall $r; $acts = @(Helper-Log $r)[$hl..99]
    Record 'A' 'successful uninstall: task, registration, add-in, versions, state, HKCU\Software\PPTK removed' ($ok0 -and $x.Code -eq 0 -and (Removed $r) -and ($acts -join ',') -eq 'stop') (Short $x)
    Record 'U' 'detect.ps1 confirms removal (NotInstalled, no leftover pieces)' (Gone $r $P100) ''
    Record 'M' 'unrelated add-in (file + registration) and unrelated task preserved' (Unrelated-Intact $r) ''
    Record 'N' 'VBA user settings preserved' ((Reg $r).$SettingsKey.OperatingMode -eq 'pro') ''
    Record 'O' 'SlideLibrary folder content and its setting preserved' ((UserData-Intact $r)) ''
    Record 'P' 'Documents backups and user presentations preserved' ((Test-Path -LiteralPath (Join-Path $r 'Documents\PPTK Flip Backups\deck (before flip).pptx')) -and (Test-Path -LiteralPath (Join-Path $r 'Documents\Quarterly review.pptx'))) ''
    $before = Snap $r; $hl = @(Helper-Log $r).Count; $x = Uninstall $r
    Record 'B' 'repeated uninstall: success, nothing changed' ($x.Code -eq 0 -and $x.Out -match 'not installed' -and (Snap $r) -eq $before -and @(Helper-Log $r).Count -eq $hl) (Short $x)
    $r = New-Root 'C'; $before = Snap $r; $x = Uninstall $r
    Record 'C' 'no installation present: success, nothing created or changed' ($x.Code -eq 0 -and (Snap $r) -eq $before -and -not (Test-Path -LiteralPath (Join-Path $r 'LocalAppData\PPTK'))) (Short $x)

    # ---------------------------------------------------------------- D, E (PowerPoint open / add-in locked)
    $r = New-Root 'D'; [void](Install $r $P100); Set-Processes $r 'POWERPNT'; $before = Snap $r
    $x = Uninstall $r
    Record 'D' 'PowerPoint open: 1618, nothing removed (registration kept), still detected' ($x.Code -eq 1618 -and $x.Err -match 'RetryLater' -and (Snap $r) -eq $before -and (Detected $r $P100)) (Short $x)
    Set-Processes $r; $x = Uninstall $r
    Record 'D2' 'PowerPoint closed: the retry removes it' ($x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100)) (Short $x)
    $r = New-Root 'E'; [void](Install $r $P100); $before = Snap $r
    $h = [IO.File]::Open((Addin $r), 'Open', 'Read', 'Read')
    try { $x = Uninstall $r } finally { $h.Dispose() }
    Record 'E' 'add-in file locked (no process listed): 1618, nothing changed, still detected' ($x.Code -eq 1618 -and (Snap $r) -eq $before -and (Detected $r $P100)) (Short $x)

    # ---------------------------------------------------------------- F, G (pending updates)
    $r = New-Root 'F'; [void](Install $r $P100); [void](Install $r $P101); Set-Processes $r 'POWERPNT'; [void](Install $r $P103)
    $before = Snap $r; $x = Uninstall $r
    Record 'F1' 'pending update + PowerPoint open: 1618, pending kept, not applied' ($x.Code -eq 1618 -and (Snap $r) -eq $before -and (Sha (Addin $r)) -eq (Sha (Join-Path $P101 'payload\PPTK.ppam')) -and (Detected $r $P101)) (Short $x)
    Set-Processes $r; $hl = @(Helper-Log $r).Count; $x = Uninstall $r; $acts = @(Helper-Log $r)[$hl..99]
    Record 'F2' 'pending update (installed 1.0.1, previous 1.0.0 intact): removed, pending discarded, never applied' ($x.Code -eq 0 -and $x.Out -match 'pending update 1\.0\.3 discarded' -and (Removed $r) -and (Gone $r $P101) -and -not (Detected $r $P103) -and ($acts -join ',') -eq 'stop') (Short $x)
    $r = New-Root 'G1'; [void](Install $r $P100); Set-Processes $r 'POWERPNT'; [void](Install $r $P101); Set-Processes $r
    [IO.File]::AppendAllText((Pptk $r 'pending\PPTK-1.0.1.ppam'), 'corrupt')
    $x = Uninstall $r
    Record 'G1' 'corrupt pending add-in (hash mismatch): removed, not applied' ($x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100)) (Short $x)
    $r = New-Root 'G2'; [void](Install $r $P100); Set-Processes $r 'POWERPNT'; [void](Install $r $P101); Set-Processes $r
    $sp = Pptk $r 'state.json'; $o = Read-Json $sp; $o.pending = [pscustomobject]@{ file = 'pending\nothing.ppam' }; Write-Json $sp $o
    $x = Uninstall $r
    Record 'G2' 'corrupt pending entry in state.json (no version): removed' ($x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100)) (Short $x)

    # ---------------------------------------------------------------- H, I (damaged managed installation)
    $r = New-Root 'H1'; [void](Install $r $P100); Remove-Item -LiteralPath (Pptk $r 'versions\1.0.0\PPTKKeys\PPTKKeys.dll'), (Addin $r)
    $x = Uninstall $r
    Record 'H1' 'missing managed files (helper DLL, add-in): removed' ($x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100)) (Short $x)
    $r = New-Root 'H2'; [void](Install $r $P100); Remove-Item -LiteralPath (Pptk $r 'versions') -Recurse
    $x = Uninstall $r
    Record 'H2' 'whole versions folder missing (no manifest): removed, verified with a minimal manifest' ($x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100)) (Short $x)
    $r = New-Root 'I'; [void](Install $r $P100); Write-Json (Join-Path $r 'tasks.json') @(Tasks $r | Where-Object { $_.name -ne 'PPTKKeys Startup' })
    $x = Uninstall $r
    Record 'I' 'missing task: removed, unrelated task kept' ($x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100) -and (Unrelated-Intact $r)) (Short $x)

    # ---------------------------------------------------------------- J, K (not verified as PPTK-owned: nothing changed)
    $refuse = { param($id, $name, $root, $pattern) $b = Snap $root; $x = Uninstall $root; Record $id $name ($x.Code -eq 1 -and $x.Err -match 'ManualApprovalRequired' -and $x.Err -match $pattern -and (Snap $root) -eq $b) (Short $x) }
    $r = New-Root 'J1'; [void](Install $r $P100); $t = @(Tasks $r); ($t | Where-Object { $_.name -eq 'PPTKKeys Startup' }).userId = 'OTHERDOM\someone'; Write-Json (Join-Path $r 'tasks.json') $t
    & $refuse 'J1' 'task owned by another user: manual approval, nothing changed' $r 'belongs to'
    $r = New-Root 'J2'; [void](Install $r $P100); $t = @(Tasks $r); ($t | Where-Object { $_.name -eq 'PPTKKeys Startup' }).actions[0].arguments = 'exec "C:\Elsewhere\PPTKKeys.dll"'; Write-Json (Join-Path $r 'tasks.json') $t
    & $refuse 'J2' 'task runs a helper outside LocalAppData\PPTK\versions: manual approval, nothing changed' $r 'PPTK-managed helper'
    $r = New-Root 'K'; [void](Install $r $P100); $reg = Reg $r; $reg.$AddinKey.Path = 'C:\Other\PPTK.ppam'; Write-Json (Join-Path $r 'registry.json') $reg
    & $refuse 'K' 'add-in registration points elsewhere: manual approval, nothing changed' $r 'registration points to'

    # ---------------------------------------------------------------- L (legacy manual installation)
    $mkLegacy = { param($name)
        $lr = New-Root $name
        [IO.File]::WriteAllText((Addin $lr), 'manually installed add-in', $Utf8)
        $lk = Pptk $lr 'PPTKKeys'; New-Item -ItemType Directory -Force -Path $lk | Out-Null
        [IO.File]::WriteAllText((Join-Path $lk 'PPTKKeys.dll'), 'manual helper', $Utf8)
        $reg = Reg $lr; $reg | Add-Member -NotePropertyName $AddinKey -NotePropertyValue ([ordered]@{ Path = 'PPTK.ppam'; AutoLoad = 1 }); Write-Json (Join-Path $lr 'registry.json') $reg
        Write-Json (Join-Path $lr 'tasks.json') (@(Tasks $lr) + @([ordered]@{ name = 'PPTKKeys Startup'; path = '\'; enabled = $true; userId = 'tester'; logonType = 'Interactive'; runLevel = 'Limited'
            triggers = @([ordered]@{ type = 'Logon'; userId = 'TESTDOM\tester'; enabled = $true })
            actions = @([ordered]@{ execute = (Join-Path $lr 'LocalAppData\Microsoft\dotnet\dotnet.exe'); arguments = 'exec "' + (Join-Path $lk 'PPTKKeys.dll') + '"' }) }))
        $lr }
    $r = & $mkLegacy 'L1'; $before = Snap $r; $x = Uninstall $r
    Record 'L1' 'manual installation: left unchanged, manual approval / migration explained' ($x.Code -eq 1 -and $x.Err -match 'ManualApprovalRequired' -and $x.Err -match 'manual installation' -and $x.Err -match 'Adopt it first' -and (Snap $r) -eq $before) (Short $x)
    $r = & $mkLegacy 'L2'; [void](Install $r $P100); $lb = Pptk $r ((State-Json $r).legacyBackup)
    $x = Uninstall $r; $y = Uninstall $r
    Record 'L2' 'adopted manual installation: uninstall keeps legacy\ backups and the old PPTKKeys folder; repeat is a no-op' ($x.Code -eq 0 -and (Removed $r) -and (Test-Path -LiteralPath (Join-Path $lb 'PPTK.ppam')) -and (Test-Path -LiteralPath (Join-Path $lb 'task.json')) -and (Test-Path -LiteralPath (Pptk $r 'PPTKKeys\PPTKKeys.dll')) -and $y.Code -eq 0 -and $y.Out -match 'not installed') ((Short $x) + ' / ' + (Short $y))

    # ---------------------------------------------------------------- Q (interrupted), R (failure + restore)
    foreach ($step in 'task', 'registration', 'addin', 'files', 'state') {
        $r = New-Root "Q-$step"; [void](Install $r $P100)
        $x = Uninstall $r @('-TestCrashAt', $step); $mid = Test-Path -LiteralPath (Pptk $r 'transaction.json'); $y = Uninstall $r
        Record "Q-$step" "uninstall killed after '$step': next run undoes, then removes completely" ($x.Code -eq 99 -and $mid -and $y.Code -eq 0 -and (Removed $r) -and (Gone $r $P100) -and (UserData-Intact $r)) ("crash exit $($x.Code); " + (Short $y))
    }
    $r = New-Root 'Q-install'; [void](Install $r $P100); $before = Snap $r -Restored
    $x = Uninstall $r @('-TestCrashAt', 'files'); $y = Install $r $P100
    Record 'Q-install' 'uninstall killed after moving files; install.ps1 runs next: everything moved back, 1.0.0 detected' ($x.Code -eq 99 -and $y.Code -eq 0 -and (Detected $r $P100) -and (Snap $r -Restored) -eq $before) (Short $y)
    foreach ($step in 'task', 'addin', 'files', 'state') {
        $r = New-Root "R-$step"; [void](Install $r $P100); $before = Snap $r -Restored
        $x = Uninstall $r @('-TestFailAt', $step)
        Record "R-$step" "failure after '$step': registration, task, files and state restored; still detected" ($x.Code -eq 1 -and $x.Err -match 'installation was restored' -and (Snap $r -Restored) -eq $before -and (Detected $r $P100) -and (State-Json $r).installed -eq '1.0.0' -and -not (Test-Path -LiteralPath (Pptk $r 'transaction.json'))) (Short $x)
    }
    $r = New-Root 'R-broken'; [void](Install $r $P100)
    $x = Uninstall $r @('-TestFailAt', 'addin+damage')
    Record 'R-broken' 'restore itself fails: clear RESTORE FAILED error, exit 1, journal kept for the next run' ($x.Code -eq 1 -and $x.Err -match 'RESTORE FAILED' -and $x.Err -notmatch 'removed' -and (Test-Path -LiteralPath (Pptk $r 'transaction.json'))) (Short $x)

    # ---------------------------------------------------------------- S (concurrency)
    $r = New-Root 'S'; [void](Install $r $P100); $before = Snap $r
    $held = [IO.File]::Open((Pptk $r 'state.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    try { $x = Uninstall $r @('-LockTimeoutSeconds', '2') } finally { $held.Dispose() }
    Record 'S1' 'uninstall while another operation holds state.lock: 1618 Busy, nothing changed' ($x.Code -eq 1618 -and $x.Err -match 'Busy' -and (Snap $r) -eq $before -and (Detected $r $P100)) (Short $x)
    # Real overlap: an install (slow: it starts detect.ps1 twice) and an uninstall started together.
    $pi = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$Install`"", '-Root', "`"$r`"", '-PackagePath', "`"$P103`"") -PassThru -WindowStyle Hidden
    Start-Sleep -Milliseconds 300
    $x = Uninstall $r @('-LockTimeoutSeconds', '60'); $pi.WaitForExit()
    $consistent = ((Removed $r) -and (Gone $r $P103) -and (Gone $r $P100)) -or (Detected $r $P103)
    Record 'S2' 'install and uninstall started together: serialized by state.lock, final state consistent' ($pi.ExitCode -in 0, 1618 -and $x.Code -in 0, 1618 -and $consistent -and -not (Test-Path -LiteralPath (Pptk $r 'transaction.json'))) ("install exit $($pi.ExitCode), uninstall " + (Short $x))

    # ---------------------------------------------------------------- MV (transient folder-move failures), TX (leftover txn\)
    $r = New-Root 'MV1'; [void](Install $r $P100); $x = Uninstall $r @('-TestMoveFailures', '2')
    $log = (Get-ChildItem -LiteralPath (Pptk $r 'logs') -Filter 'uninstall-*' -File | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
    Record 'MV1' 'moving files aside fails twice (transient): retried, removed' ($x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100) -and $log -match 'attempt 2 failed') (Short $x)
    $r = New-Root 'MV2'; [void](Install $r $P100); $x = Uninstall $r @('-TestMoveFailures', '5')
    Record 'MV2' 'moving files aside keeps failing (5): exit 1, installation restored and detected' ($x.Code -eq 1 -and $x.Err -match 'installation was restored' -and (Detected $r $P100) -and -not (Test-Path -LiteralPath (Pptk $r 'transaction.json'))) (Short $x)
    $r = New-Root 'TX'; [void](Install $r $P100); New-Item -ItemType Directory -Force -Path (Pptk $r 'txn\removed\versions') | Out-Null
    $x = Uninstall $r
    Record 'TX' 'leftover txn\ without a journal: removed first, then uninstall completes' ($x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100)) (Short $x)

    # ---------------------------------------------------------------- RU (file replace retry)
    $r = New-Root 'RU1'; [void](Install $r $P100); $x = Uninstall $r @('-TestReplaceFailures', '3')
    Record 'RU1' 'file replacement fails 3 times (transient): retried, removed' ($x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100)) (Short $x)
    $r = New-Root 'RU2'; [void](Install $r $P100); $x = Uninstall $r @('-TestReplaceFailures', '5')
    Record 'RU2' 'file replacement keeps failing (5): exit 1, still installed and detected, no journal' ($x.Code -eq 1 -and (Detected $r $P100) -and -not (Test-Path -LiteralPath (Pptk $r 'transaction.json'))) (Short $x)

    # ---------------------------------------------------------------- T (Arabic / Unicode)
    $arabic = -join (@(0x0627, 0x062E, 0x062A, 0x0628, 0x0627, 0x0631, 0x0020, 0x2713) | ForEach-Object { [char]$_ })
    $r = New-Root $arabic; $a = Install $r $P100; $x = Uninstall $r
    Record 'T' 'Arabic/Unicode root: install then uninstall, removal confirmed, user data kept' ($a.Code -eq 0 -and $x.Code -eq 0 -and (Removed $r) -and (Gone $r $P100) -and (UserData-Intact $r)) ("$($a.Code)/$($x.Code)")

    # ---------------------------------------------------------------- V (nothing outside -Root)
    $code = @([IO.File]::ReadAllLines($Uninstall) | Where-Object { $_ -notmatch '^\s*#' })
    $forbidden = @($code | Where-Object { $_ -match '(?i)Register-ScheduledTask|Set-ScheduledTask|Start-ScheduledTask|Stop-ScheduledTask|Unregister-ScheduledTask|Get-ScheduledTask|Stop-Process|Start-Process|Get-Process|Set-ItemProperty|New-ItemProperty|Remove-ItemProperty|Remove-ItemProperty|HKCU:|HKEY_CURRENT_USER|Registry::|\$env:(APPDATA|LOCALAPPDATA|USERPROFILE)|GetFolderPath' })
    Record 'V1' 'uninstall.ps1 has no live registry, task, process or profile operations' ($forbidden.Count -eq 0) ($forbidden -join ' / ')
    $liveAfter = Live-Snap
    $diff = @('Task', 'PptkKey', 'AddinReg', 'Addin', 'Local') | Where-Object { $liveBefore.$_ -ne $liveAfter.$_ }
    Record 'V2' 'live task, HKCU, AddIns file and LocalAppData\PPTK unchanged' ($diff.Count -eq 0) ("changed: " + ($diff -join ', '))
    Record 'V3' 'running PPTKKeys and PowerPoint processes unchanged' ($liveBefore.Helper -eq $liveAfter.Helper -and $liveBefore.PowerPoint -eq $liveAfter.PowerPoint) ("helper pid " + $liveAfter.Helper)
} catch {
    Record 'ERR' 'test run' $false ($_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    Remove-Item -LiteralPath $Base -Recurse -Force -ErrorAction SilentlyContinue
}

$failed = @($results | Where-Object { -not $_.Ok })
Write-Host ("TOTAL {0} tests, {1} failed" -f $results.Count, $failed.Count)
if ($failed.Count) { exit 1 } else { exit 0 }
