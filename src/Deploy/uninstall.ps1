<#
.SYNOPSIS
  PPTK per-user uninstaller: removes a managed PPTK installation.

.DESCRIPTION
  Production (no -Root): removes PPTK for the signed-in user who runs it (Intune install behavior:
  User). Refused for SYSTEM. The running PPTKKeys helper is asked to close (never terminated); if it
  does not close, nothing is changed and the exit code is 1618.

  Test modes (src\Deploy\deploy-env.ps1): -Root <test root> simulates everything (registry.json,
  tasks.json, processes.json, AppData, LocalAppData, helper-actions.log); -Root with -LiveTestId
  <12 hex> uses the real Task Scheduler and HKCU under PPTK-Test-<id> names only. A test root must
  contain PPTK-TEST-ROOT; otherwise exit 2, nothing changed.

  Removed (only when verified as PPTK-managed):
    Task \PPTKKeys Startup (current user, dotnet exec "...LocalAppData\PPTK\versions\<v>\PPTKKeys\PPTKKeys.dll")
    HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK  (Path = PPTK.ppam)
    AppData\Microsoft\AddIns\PPTK.ppam
    LocalAppData\PPTK\versions\, pending\, state.json, HKCU\Software\PPTK
  Kept: user settings (VBA registry), Documents, the SlideLibrary folder, other add-ins,
    LocalAppData\PPTK\legacy\ (backups of a manual installation), LocalAppData\PPTK\logs\,
    a manual installation's LocalAppData\PPTK\PPTKKeys\ folder, and anything else.

  A pending update is removed with everything else; it is never applied.
  Files are not deleted while the operation runs: they are moved into LocalAppData\PPTK\txn\ and
  deleted only after detect.ps1 confirms the removal, so a failure or an interrupted run puts
  everything back (transaction.json, same journal as install.ps1, same state.lock).

  Exit codes:
    0     removed and confirmed by detect.ps1, or nothing installed (no change)
    1618  retry later: PowerPoint is running or has the add-in open, the helper did not close, or
          another PPTK operation holds state.lock. Nothing was changed.
    1     failure (the installation is restored) or manual approval required (nothing changed)
    2     unsafe test root; nothing changed

.PARAMETER TestMoveFailures / TestReplaceFailures
  Test hooks: the first N folder moves / file replacements fail with a simulated transient IOException
  (retried up to 5 times).

.PARAMETER TestFailAt / TestCrashAt
  Test hooks. Steps: task, registration, addin, files, state.
#>
[CmdletBinding()]
param(
    [string]$Root = '',
    [string]$LiveTestId = '',
    [int]$LockTimeoutSeconds = 30,
    [string]$TestFailAt = '',
    [string]$TestCrashAt = '',
    [int]$TestMoveFailures = 0,
    [int]$TestReplaceFailures = 0
)

$ErrorActionPreference = 'Stop'

# Intune may start a 32-bit PowerShell. Task Scheduler, Program Files and the .NET host must be the 64-bit
# ones, so run again in the 64-bit PowerShell with the same parameters and return its exit code.
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $ps64 = Join-Path $env:windir 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    $argv = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    foreach ($k in $PSBoundParameters.Keys) { $v = $PSBoundParameters[$k]; if ($v -is [Management.Automation.SwitchParameter]) { if ($v.IsPresent) { $argv += "-$k" } } else { $argv += "-$k"; $argv += [string]$v } }
    & $ps64 @argv
    exit $LASTEXITCODE
}
$KeysFiles = @('PPTKKeys.dll', 'PPTKKeys.deps.json', 'PPTKKeys.runtimeconfig.json')
$ManagedItems = @('versions', 'pending')
$Utf8 = New-Object Text.UTF8Encoding($false)

class InstallStop : Exception { [int]$Code; InstallStop([int]$code, [string]$msg) : base($msg) { $this.Code = $code } }
function Stop-Uninstall([int]$code, [string]$category, [string]$detail) { throw [InstallStop]::new($code, "PPTK-UNINSTALL ${category}: $detail") }

# ---------------------------------------------------------------- helpers (same as install.ps1)
function Log([string]$m) { if ($script:LogPath) { [IO.File]::AppendAllText($script:LogPath, ([DateTime]::UtcNow.ToString('o') + ' ' + $m + "`r`n"), $Utf8) } }
function Has([object]$o, [string]$n) { return ($null -ne $o) -and ($o.PSObject.Properties.Name -contains $n) }
function Test-SemVer([object]$v) { return ($v -is [string]) -and ($v -match '^\d+\.\d+\.\d+$') }
function Get-Sha256([string]$path) {
    $s = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    $h = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($h.ComputeHash($s)).Replace('-', '')).ToLowerInvariant() } finally { $h.Dispose(); $s.Dispose() }
}
function Same-Path([string]$a, [string]$b) { if (-not $a -or -not $b) { return $false }; return [string]::Equals([IO.Path]::GetFullPath($a).TrimEnd('\'), [IO.Path]::GetFullPath($b).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) }
function Test-Under([string]$path, [string]$dir) { $p = [IO.Path]::GetFullPath($path); $d = [IO.Path]::GetFullPath($dir).TrimEnd('\') + '\'; return $p.StartsWith($d, [StringComparison]::OrdinalIgnoreCase) }
function ConvertTo-JsonText($obj) {
    if ($obj -is [Array]) { return '[' + ((@($obj) | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 10 }) -join ',') + ']' }
    return (ConvertTo-Json -InputObject $obj -Depth 10)
}
# ---- bounded retry for file replace/move (same rules as install.ps1): transient sharing/lock and
# replace failures (32, 33, 1175, 1176) 5 attempts; access denied 3, none on a read-only target; others 1.
function Get-IoAttempts($e, [string]$target) {
    $code = $e.HResult -band 0xFFFF
    if ($e -is [IO.IOException] -and $code -in 32, 33, 1175, 1176) { return 5 }
    if ($e -is [UnauthorizedAccessException] -or ($e -is [IO.IOException] -and $code -eq 5)) {
        if ($target -and (Test-Path -LiteralPath $target -PathType Leaf) -and ((Get-Item -LiteralPath $target -Force).Attributes -band [IO.FileAttributes]::ReadOnly)) { return 1 }
        return 3
    }
    return 1
}
function Invoke-IoRetry([string]$ioWhat, [string]$ioTarget, [string]$ioHook, [scriptblock]$ioOp, [scriptblock]$ioDone, [scriptblock]$ioGuard) {
    for ($ioAttempt = 1; ; $ioAttempt++) {
        if ($ioAttempt -gt 1) {
            if ($ioDone -and (& $ioDone)) { Log "${ioWhat}: the earlier attempt had completed"; return }
            if ($ioGuard) { & $ioGuard }
        }
        try {
            Test-IoHook $ioHook $ioTarget
            & $ioOp
            return
        } catch {
            $ioErr = $_.Exception; while ($ioErr -is [Management.Automation.MethodInvocationException] -and $ioErr.InnerException) { $ioErr = $ioErr.InnerException }
            $ioMax = Get-IoAttempts $ioErr $ioTarget
            if ($ioAttempt -ge $ioMax) { if ($ioMax -gt 1) { Log "${ioWhat}: attempt $ioAttempt of $ioMax failed; giving up" }; throw }
            Log "${ioWhat}: attempt $ioAttempt failed (0x$('{0:X8}' -f $ioErr.HResult) $($ioErr.Message)); retrying"
            Start-Sleep -Milliseconds (200 * $ioAttempt)
        }
    }
}
function Test-IoHook([string]$hook, [string]$target) {
    if ($hook -eq 'move' -and $script:TestMoveFailures -gt 0) { $script:TestMoveFailures--; throw [IO.IOException]::new("simulated transient failure moving $target", -2147024864) }
    if ($hook -eq 'replace' -and $script:TestReplaceFailures -gt 0) { $script:TestReplaceFailures--; throw [IO.IOException]::new("simulated: Unable to remove the file to be replaced ($target)", -2147023721) }
}

function Write-Atomic([string]$path, [string]$text) {
    $dir = Split-Path -Parent $path; if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $tmp = "$path.tmp-" + [guid]::NewGuid().ToString('N').Substring(0, 8)
    [IO.File]::WriteAllText($tmp, $text, $Utf8)
    Invoke-IoRetry "replace $([IO.Path]::GetFileName($path))" $path 'replace' `
        { if (Test-Path -LiteralPath $path) { [IO.File]::Replace($tmp, $path, [NullString]::Value) } else { [IO.File]::Move($tmp, $path) } } `
        { -not (Test-Path -LiteralPath $tmp) -and (Test-Path -LiteralPath $path) -and [IO.File]::ReadAllText($path) -ceq $text } `
        { if (-not (Test-Path -LiteralPath $tmp)) { throw "the temporary file for $path disappeared" } }
}
function Read-Json([string]$path) { if (Test-Path -LiteralPath $path) { return ([IO.File]::ReadAllText($path) | ConvertFrom-Json) } else { return $null } }
# ---------------------------------------------------------------- environment (production / sim / live-test)
$EnvPrefix = 'PPTK-UNINSTALL'
. (Join-Path $PSScriptRoot 'deploy-env.ps1')
$PptkDir = Join-Path $LocalAppData 'PPTK'
$AddinFile = Join-Path $AppData 'Microsoft\AddIns\PPTK.ppam'
$StatePath = Join-Path $PptkDir 'state.json'
$JournalPath = Join-Path $PptkDir 'transaction.json'
$TxnDir = Join-Path $PptkDir 'txn'
$DetectScript = Join-Path $PSScriptRoot 'detect.ps1'
$script:LogPath = $null

function Test-IsCurrentUser([string]$v) {
    if (-not $v) { return $false }
    if ($v -ieq $Identity.sid -or $v -ieq $Identity.name) { return $true }
    return ($v -notmatch '\\') -and ($v -ieq (($Identity.name -split '\\')[-1]))
}
# Folder moves (same bounded retry; a move that completed is not repeated).
function Move-WithRetry([string]$from, [string]$to) {
    Invoke-IoRetry "move $([IO.Path]::GetFileName($from))" $from 'move' `
        { Move-Item -LiteralPath $from -Destination $to -ErrorAction Stop } `
        { -not (Test-Path -LiteralPath $from) -and (Test-Path -LiteralPath $to) } `
        { if (-not (Test-Path -LiteralPath $from)) { throw "$from disappeared before it could be moved" } }
}

function Step-Hook([string]$step) {
    if ($TestCrashAt -eq $step) { Log "TEST: simulated crash after step $step"; [Environment]::Exit(99) }
    if ($TestFailAt -eq $step) { throw "simulated failure after step $step" }
    # "<step>+damage": also damage the journal's add-in copy, so the undo itself fails (report test).
    if ($TestFailAt -eq "$step+damage") { [IO.File]::AppendAllText((Join-Path $TxnDir 'PPTK.ppam.before'), 'x'); throw "simulated failure after step $step (journal damaged)" }
}

# ---------------------------------------------------------------- current installation
function Get-Installation {
    $tasks = Get-Tasks
    $main = @($tasks | Where-Object { $_.name -eq $TaskName -and $_.path -eq $TaskPath })
    $dups = @($tasks | Where-Object { -not ($_.name -eq $TaskName -and $_.path -eq $TaskPath) -and @(@($_.actions) | Where-Object { "$($_.execute) $($_.arguments)" -match 'PPTKKeys\.dll' }).Count })
    $state = $null; $stateCorrupt = $false
    if (Test-Path -LiteralPath $StatePath) { try { $state = [IO.File]::ReadAllText($StatePath) | ConvertFrom-Json; if (-not (Has $state 'installed')) { $stateCorrupt = $true } } catch { $stateCorrupt = $true } }
    $regVersion = Get-Reg $PptkKey 'InstalledVersion'
    $version = if ($state -and -not $stateCorrupt -and $state.installed) { [string]$state.installed } elseif ($regVersion) { [string]$regVersion } else { $null }
    $managed = [bool]($regVersion -or (Test-Path -LiteralPath $StatePath))
    $approval = @()
    if ($managed) {
        if ($version -and -not (Test-SemVer $version)) { $approval += "installed version '$version' is not a version" }
        if ($dups.Count) { $approval += ('other tasks run PPTKKeys: ' + (($dups | ForEach-Object { "$($_.path)$($_.name)" }) -join ', ')) }
        if ($main.Count) {
            $t = $main[0]; $a = @($t.actions)
            if (-not (Test-IsCurrentUser $t.userId)) { $approval += "task '$TaskName' belongs to '$($t.userId)', not the current user" }
            $m0 = if ($a.Count -eq 1) { [regex]::Match([string]$a[0].arguments, '^\s*exec\s+"([^"]+)"\s*$') } else { $null }
            if (-not $m0 -or -not $m0.Success -or ([IO.Path]::GetFileName(([string]$a[0].execute).Trim('"'))) -ine 'dotnet.exe' -or -not (Test-Under $m0.Groups[1].Value (Join-Path $PptkDir 'versions')) -or [IO.Path]::GetFileName($m0.Groups[1].Value) -ine 'PPTKKeys.dll') {
                $approval += "task '$TaskName' does not run a PPTK-managed helper (dotnet.exe exec ""LocalAppData\PPTK\versions\<v>\PPTKKeys\PPTKKeys.dll"")"
            }
        }
        if (Test-RegKey $AddinKey) {
            $rp = [string](Get-Reg $AddinKey 'Path')
            if (-not ($rp -ieq 'PPTK.ppam' -or (Same-Path $rp $AddinFile))) { $approval += "add-in registration points to '$rp', not to the PPTK add-in" }
        }
    }
    $legacy = @()
    if (-not $managed) {
        if (Test-Path -LiteralPath $AddinFile) { $legacy += 'AddIns\PPTK.ppam' }
        if (Test-RegKey $AddinKey) { $legacy += 'add-in registration' }
        if ($main.Count -or $dups.Count) { $legacy += 'PPTKKeys task' }
        # A leftover LocalAppData\PPTK\PPTKKeys folder alone (kept after adoption) runs nothing: not an installation.
    }
    return [pscustomobject]@{ Managed = $managed; Version = $version; State = $state; Task = $(if ($main.Count) { $main[0] } else { $null }); Approval = $approval; Legacy = $legacy }
}

# ---------------------------------------------------------------- journal (same file and fields as install.ps1)
function Save-Journal([string]$version) {
    New-Item -ItemType Directory -Force -Path $TxnDir | Out-Null
    $addinBackup = $null
    if (Test-Path -LiteralPath $AddinFile) { $addinBackup = Join-Path $TxnDir 'PPTK.ppam.before'; Copy-Item -LiteralPath $AddinFile -Destination $addinBackup -Force }
    $j = [ordered]@{
        journalVersion = 1; tasksFormat = $TasksFormat   # src\Deploy\JOURNAL.md
        operation = 'uninstall'; target = $version; startedUtc = [DateTime]::UtcNow.ToString('o')
        stateBefore = $(if (Test-Path -LiteralPath $StatePath) { [IO.File]::ReadAllText($StatePath) } else { $null })
        tasksBefore = (Export-Tasks)
        pptkKeyBefore = (Export-RegKey $PptkKey)
        addinKeyBefore = (Export-RegKey $AddinKey)
        addinBefore = $addinBackup; addinBeforeSha = $(if ($addinBackup) { Get-Sha256 $addinBackup } else { $null })
        createdVersionDir = $null
        # Folders moved into txn\removed\ (relative to LocalAppData\PPTK); moved back by an undo.
        moved = @($ManagedItems | Where-Object { Test-Path -LiteralPath (Join-Path $PptkDir $_) } | ForEach-Object { [ordered]@{ from = $_; to = "txn\removed\$_" } })
    }
    Write-Atomic $JournalPath (ConvertTo-JsonText $j)
    Log "journal written for uninstall $version"
}

# Puts everything back as recorded before the operation (also undoes an interrupted install.ps1 run).
function Undo-Journal([string]$why) {
    $j = Read-Json $JournalPath
    if (-not $j) { return }
    Log "undo ($why): restoring the state from before $($j.operation) $($j.target)"
    Assert-JournalFormat $j
    Stop-HelperForUndo
    foreach ($mv in @($j.moved | Where-Object { $_ })) {
        $src = Join-Path $PptkDir $mv.to; $dst = Join-Path $PptkDir $mv.from
        if ((Test-Path -LiteralPath $src) -and -not (Test-Path -LiteralPath $dst)) { Move-WithRetry $src $dst }
        if (-not (Test-Path -LiteralPath $dst)) { throw "$($mv.from) could not be restored" }
    }
    if ($j.addinBefore) {
        if ((Get-Sha256 $j.addinBefore) -ne $j.addinBeforeSha) { throw 'the add-in backup in the journal is damaged' }
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $AddinFile) | Out-Null
        if (-not (Test-Path -LiteralPath $AddinFile) -or (Get-Sha256 $AddinFile) -ne $j.addinBeforeSha) { Copy-Item -LiteralPath $j.addinBefore -Destination $AddinFile -Force }
        if ((Get-Sha256 $AddinFile) -ne $j.addinBeforeSha) { throw 'the add-in could not be restored' }
    } elseif (Test-Path -LiteralPath $AddinFile) { Remove-Item -LiteralPath $AddinFile -Force }
    # tasksBefore = null: there was no PPTKKeys task; remove only PPTKKeys tasks, keep every other task.
    Import-Tasks $j.tasksBefore
    if ($null -ne $j.pptkKeyBefore) { Set-RegKey $PptkKey ($j.pptkKeyBefore | ConvertFrom-Json) } else { Set-RegKey $PptkKey $null }
    if ($null -ne $j.addinKeyBefore) { Set-RegKey $AddinKey ($j.addinKeyBefore | ConvertFrom-Json) } else { Set-RegKey $AddinKey $null }
    if ($null -ne $j.stateBefore) { Write-Atomic $StatePath $j.stateBefore } elseif (Test-Path -LiteralPath $StatePath) { Remove-Item -LiteralPath $StatePath -Force }
    Remove-CreatedVersionDir $j.createdVersionDir
    # The journal first (a kept journal would be undone again); txn\ is best effort.
    Remove-Item -LiteralPath $JournalPath -Force
    Remove-Item -LiteralPath $TxnDir -Recurse -Force -ErrorAction SilentlyContinue
    [void](Helper-Action 'restart (restored)')
    Log 'undo complete'
}

function Invoke-Detect([string]$manifest) {
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$DetectScript`" -ManifestPath `"$manifest`"" + $(if ($Root) { " -Root `"$Root`"" } else { '' }) + $(if ($LiveTestId) { " -LiveTestId $LiveTestId" } else { '' })
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8; $psi.StandardErrorEncoding = [Text.Encoding]::UTF8
    $p = [Diagnostics.Process]::Start($psi); $o = $p.StandardOutput.ReadToEnd(); $e = $p.StandardError.ReadToEnd(); $p.WaitForExit()
    return [pscustomobject]@{ Code = $p.ExitCode; Out = $o.Trim(); Err = $e.Trim() }
}

# The manifest detect.ps1 checks against: the installed version's own version.json (copied before
# the folders move), or, when it is missing, a minimal one with that version number.
function Save-DetectManifest([string]$version) {
    $dest = Join-Path $TxnDir 'detect-manifest.json'
    $own = Join-Path $PptkDir "versions\$version\version.json"
    if ($version -and (Test-Path -LiteralPath $own)) { Copy-Item -LiteralPath $own -Destination $dest -Force; return $dest }
    $v = if (Test-SemVer $version) { $version } else { '0.0.0' }
    $zero = '0' * 64
    $files = @(@('payload/PPTK.ppam') + @($KeysFiles | ForEach-Object { "payload/PPTKKeys/$_" }) | ForEach-Object { [ordered]@{ path = $_; sha256 = $zero } })
    [IO.File]::WriteAllText($dest, (ConvertTo-JsonText ([ordered]@{ schemaVersion = 1; packageName = 'PPTK'; packageVersion = $v; files = $files })), $Utf8)
    return $dest
}

# ---------------------------------------------------------------- operation
function Invoke-Uninstall {
    $cur = Get-Installation
    if (-not $cur.Managed) {
        if ($cur.Legacy.Count) {
            Stop-Uninstall 1 'ManualApprovalRequired' ("PPTK is not installed by the managed installer; found a manual installation (" + ($cur.Legacy -join ', ') + "). Nothing was changed. Adopt it first with install.ps1 (it backs up the manual add-in and task to LocalAppData\PPTK\legacy\), or remove it by hand after approval.")
        }
        Log 'PPTK is not installed; nothing to do'
        return 'PPTK is not installed (no change)'
    }
    if ($cur.Approval.Count) { Stop-Uninstall 1 'ManualApprovalRequired' (($cur.Approval -join '; ') + '. Nothing was changed.') }
    # PowerPoint holds (or may rewrite) the add-in and its registration: remove nothing while it runs.
    if (Test-PowerPointRunning) { Stop-Uninstall 1618 'RetryLater' 'PowerPoint is running; close PowerPoint and run the uninstall again (nothing was changed)' }
    if (Test-Path -LiteralPath $AddinFile) {
        try { $s = [IO.File]::Open($AddinFile, 'Open', 'ReadWrite', 'None'); $s.Dispose() } catch { Stop-Uninstall 1618 'RetryLater' 'the add-in file is in use; close PowerPoint and run the uninstall again (nothing was changed)' }
    }
    $pend = if ($cur.State -and (Has $cur.State 'pending') -and $cur.State.pending) { $pv = if ((Has $cur.State.pending 'version') -and (Test-SemVer $cur.State.pending.version)) { "pending update $($cur.State.pending.version)" } else { 'corrupt pending update' }; " ($pv discarded, not applied)" } elseif (Test-Path -LiteralPath (Join-Path $PptkDir 'pending')) { ' (pending files discarded, not applied)' } else { '' }
    Log "uninstalling PPTK $($cur.Version)$pend"
    # The helper runs from versions\: it must close before its folder can be removed. It is asked to close,
    # never terminated; if it does not close, it is started again and nothing is changed.
    $helperRuns = $cur.Task -or ($Mode -eq 'production' -and @(try { Get-HelperWindows } catch { 'unknown' }).Count)
    if ($helperRuns -and -not (Helper-Action 'stop')) { [void](Helper-Action 'start'); Stop-Uninstall 1618 'RetryLater' 'the PPTKKeys helper did not close; run the uninstall again later (nothing was changed)' }

    Save-Journal $cur.Version
    $manifest = Save-DetectManifest $cur.Version

    if ($cur.Task) { Remove-MainTask; Log "task $TaskPath$TaskName removed" }
    Step-Hook 'task'
    if (Test-RegKey $AddinKey) { Set-RegKey $AddinKey $null; Log 'add-in registration removed' }
    Step-Hook 'registration'
    if (Test-Path -LiteralPath $AddinFile) { Remove-Item -LiteralPath $AddinFile -Force; Log 'AddIns\PPTK.ppam removed (copy kept in the journal until verified)' }
    Step-Hook 'addin'
    New-Item -ItemType Directory -Force -Path (Join-Path $TxnDir 'removed') | Out-Null
    foreach ($mv in @((Read-Json $JournalPath).moved | Where-Object { $_ })) { Move-WithRetry (Join-Path $PptkDir $mv.from) (Join-Path $PptkDir $mv.to); Log "$($mv.from) moved aside" }
    Step-Hook 'files'
    if (Test-Path -LiteralPath $StatePath) { Remove-Item -LiteralPath $StatePath -Force }
    if (Test-RegKey $PptkKey) { Set-RegKey $PptkKey $null }
    Step-Hook 'state'

    # Verified removal: detect.ps1 must say NotInstalled, and nothing PPTK-managed may be left.
    $det = Invoke-Detect $manifest
    if ($det.Code -eq 0 -or $det.Err -notmatch 'NotInstalled') { throw "detect.ps1 does not confirm the removal: $($det.Out) $($det.Err)" }
    $left = @()
    if (@(Get-Tasks | Where-Object { ($_.name -eq $TaskName -and $_.path -eq $TaskPath) -or @(@($_.actions) | Where-Object { "$($_.execute) $($_.arguments)" -match 'PPTKKeys\.dll' }).Count }).Count) { $left += 'PPTKKeys task' }
    if (Test-RegKey $AddinKey) { $left += 'add-in registration' }
    if (Test-Path -LiteralPath $AddinFile) { $left += 'AddIns\PPTK.ppam' }
    foreach ($i in $ManagedItems + 'state.json') { if (Test-Path -LiteralPath (Join-Path $PptkDir $i)) { $left += "LocalAppData\PPTK\$i" } }
    if (Test-RegKey $PptkKey) { $left += 'HKCU\Software\PPTK' }
    if ($left.Count) { throw ('still present after removal: ' + ($left -join ', ')) }

    # Verified. Deleting the journal is the commit; then the moved folders and the add-in copy (only txn\,
    # which this run created) are deleted best effort: a failure there never restores the removed files.
    Remove-Item -LiteralPath $JournalPath -Force
    Remove-Item -LiteralPath $TxnDir -Recurse -Force -ErrorAction SilentlyContinue
    $leftover = if (Test-Path -LiteralPath $TxnDir) { ' (some removed files are deleted on the next run)' } else { '' }
    return "PPTK $($cur.Version) removed$pend; detect.ps1 reports NotInstalled$leftover"
}

# ---------------------------------------------------------------- main
$lock = $null; $code = 0
try {
    # Nothing under LocalAppData\PPTK: no managed installation (and nothing to lock or log).
    if (-not (Test-Path -LiteralPath $PptkDir)) {
        $msg = Invoke-Uninstall
        [Console]::Out.WriteLine("PPTK-UNINSTALL OK: $msg")
        exit 0
    }
    New-Item -ItemType Directory -Force -Path (Join-Path $PptkDir 'logs') | Out-Null
    $script:LogPath = Join-Path $PptkDir ('logs\uninstall-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + $PID + '.log')
    Log 'start: uninstall'

    # state.lock: the same exclusive lock install.ps1 holds, so install and uninstall never overlap.
    $deadline = [DateTime]::UtcNow.AddSeconds($LockTimeoutSeconds)
    while (-not $lock) {
        try { $lock = [IO.File]::Open((Join-Path $PptkDir 'state.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch { if ([DateTime]::UtcNow -gt $deadline) { Stop-Uninstall 1618 'Busy' 'another PPTK operation holds state.lock; nothing was changed' }; Start-Sleep -Milliseconds 250 }
    }

    if (Test-Path -LiteralPath $JournalPath) { Undo-Journal 'an earlier operation was interrupted' }
    else { Remove-Item -LiteralPath $TxnDir -Recurse -Force -ErrorAction SilentlyContinue }   # left by a committed run

    $msg = Invoke-Uninstall
    Log "success: $msg"
    [Console]::Out.WriteLine("PPTK-UNINSTALL OK: $msg")
} catch {
    $ex = $_.Exception; if ($ex -is [Management.Automation.MethodInvocationException] -and $ex.InnerException) { $ex = $ex.InnerException }
    if ($ex -is [InstallStop]) { $code = $ex.Code; $text = $ex.Message }
    else {
        $code = 1; $text = "PPTK-UNINSTALL Failed: $($ex.Message)"
        try { if (Test-Path -LiteralPath $JournalPath) { Undo-Journal 'the uninstall failed'; $text += ' | the installation was restored' } } catch { $text += " | RESTORE FAILED: $($_.Exception.Message) (transaction.json kept for the next run)" }
        try { $st = Read-Json $StatePath; if ($st) { $st | Add-Member -NotePropertyName lastError -NotePropertyValue $ex.Message -Force; Write-Atomic $StatePath (ConvertTo-JsonText $st) } } catch { }
    }
    Log $text
    [Console]::Error.WriteLine($text)
} finally {
    if ($lock) { $lock.Dispose() }
}
exit $code
