<#
.SYNOPSIS
  PPTK per-user installer: install, upgrade, legacy adoption, pending add-in, rollback.

.DESCRIPTION
  Production (no -Root): installs for the signed-in user who runs it (Intune install behavior: User).
  Refused for SYSTEM. Run from the package folder: powershell.exe -NoProfile -ExecutionPolicy Bypass
  -File install.ps1   (the package is the folder of this script unless -PackagePath is given).

  Test modes (src\Deploy\deploy-env.ps1): -Root <test root> simulates everything; -Root with
  -LiveTestId <12 hex> uses the real Task Scheduler and HKCU under PPTK-Test-<id> names only.
  A test root must contain the marker file PPTK-TEST-ROOT; otherwise exit 2, nothing changed.
  Simulated environment files (sim mode; detect.ps1 reads the same):

    <Root>\PPTK-TEST-ROOT      marker (required)
    <Root>\identity.json       {"name": "DOMAIN\\user", "sid": "S-..."}
    <Root>\registry.json       HKCU values: {"Software\\PPTK": {...}, "<subkey>": {...}, ...}
    <Root>\tasks.json          scheduled tasks (detect.ps1 schema)
    <Root>\processes.json      running process names, e.g. ["POWERPNT"] (PowerPoint open)
    <Root>\AppData\...         %APPDATA%      <Root>\LocalAppData\...  %LOCALAPPDATA%
    <Root>\ProgramFiles\...    %ProgramFiles% (machine-wide .NET)
    <Root>\helper-actions.log  simulated stop/start of the PPTKKeys task

  Layout written (approved plan):
    LocalAppData\PPTK\versions\<v>\{PPTK.ppam, version.json, PPTKKeys\...}   kept: installed + previous
    LocalAppData\PPTK\pending\PPTK-<v>.ppam      add-in staged while PowerPoint has the old one open
    LocalAppData\PPTK\state.json                 {installed, previous, rolledBackFrom, pending, ...}
    LocalAppData\PPTK\state.lock                 held open exclusively while an operation runs
    LocalAppData\PPTK\transaction.json           undo journal of the running operation
    LocalAppData\PPTK\logs\install-*.log
    AppData\Microsoft\AddIns\PPTK.ppam
    HKCU\Software\PPTK  InstalledVersion
    HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK  Path=PPTK.ppam, AutoLoad=1
    Task \PPTKKeys Startup: current user, interactive, limited, logon trigger,
         "<dotnet.exe>" exec "LocalAppData\PPTK\versions\<v>\PPTKKeys\PPTKKeys.dll"

  Exit codes:
    0     success; detect.ps1 reports the version installed
    1618  retry later: the add-in is staged in pending\ (PowerPoint has the old one open), or
          another PPTK operation holds state.lock. The installed version is unchanged.
    1     failure; the previous installation is restored (or was never changed)
    2     unsafe test root; nothing changed

.PARAMETER Rollback
  With an update pending: cancel it and keep the installed version. Otherwise restore the previous
  verified version recorded in state.json (files from versions\<previous>); state.json then records
  rolledBackFrom and no previous version, so a second rollback is refused instead of going forward.

.PARAMETER TestMoveFailures / TestReplaceFailures / TestReplaceTarget
  Test hooks: the first N folder moves / file replacements (optionally only of the file named by
  TestReplaceTarget) fail with a simulated transient IOException; they are retried up to 5 times.

.PARAMETER TestFailAt / TestCrashAt
  Test hooks. FailAt throws at a step (the operation undoes itself); CrashAt ends the process at a
  step without any clean-up (the next run finds transaction.json and undoes the interrupted work).
  Steps: staged, task, registration, addin, state.
#>
[CmdletBinding()]
param(
    [string]$Root = '',
    [string]$LiveTestId = '',
    [string]$PackagePath = '',
    [switch]$Rollback,
    [int]$LockTimeoutSeconds = 30,
    [string]$TestFailAt = '',
    [string]$TestCrashAt = '',
    [int]$TestMoveFailures = 0,
    [int]$TestReplaceFailures = 0,
    [string]$TestReplaceTarget = ''
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
$BannedExtensions = '\.(exe|pdb|msi|bat|cmd|com|scr|vbs|js|ps1xml|lnk)$'
$RuntimePath = '^payload/runtime/(win-x64|win-arm64)/'   # the only place a package may carry executables (official .NET)
$Utf8 = New-Object Text.UTF8Encoding($false)

class InstallStop : Exception { [int]$Code; InstallStop([int]$code, [string]$msg) : base($msg) { $this.Code = $code } }
function Stop-Install([int]$code, [string]$category, [string]$detail) { throw [InstallStop]::new($code, "PPTK-INSTALL ${category}: $detail") }

# ---------------------------------------------------------------- helpers
function Log([string]$m) { if ($script:LogPath) { [IO.File]::AppendAllText($script:LogPath, ([DateTime]::UtcNow.ToString('o') + ' ' + $m + "`r`n"), $Utf8) } }
function Has([object]$o, [string]$n) { return ($null -ne $o) -and ($o.PSObject.Properties.Name -contains $n) }
function Test-SemVer([object]$v) { return ($v -is [string]) -and ($v -match '^\d+\.\d+\.\d+$') }
function Get-Sha256([string]$path) {
    $s = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    $h = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($h.ComputeHash($s)).Replace('-', '')).ToLowerInvariant() } finally { $h.Dispose(); $s.Dispose() }
}
function Same-Path([string]$a, [string]$b) { if (-not $a -or -not $b) { return $false }; return [string]::Equals([IO.Path]::GetFullPath($a).TrimEnd('\'), [IO.Path]::GetFullPath($b).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) }

function ConvertTo-JsonText($obj) {
    # Windows PowerShell 5.1 writes an array that came from ConvertFrom-Json as {"value": [...]}.
    if ($obj -is [Array]) { return '[' + ((@($obj) | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 10 }) -join ',') + ']' }
    return (ConvertTo-Json -InputObject $obj -Depth 10)
}
# ---- bounded retry for file replace/move (transient failures seen under load: another process briefly has a
# just-written file open). Attempts by failure kind; anything else fails at once:
#   sharing/lock violation, "unable to remove/move" during replace (32, 33, 1175, 1176): 5 attempts
#   access denied: 3 attempts, none when the target is read-only (a permanent condition)
function Get-IoAttempts($e, [string]$target) {
    $code = $e.HResult -band 0xFFFF
    if ($e -is [IO.IOException] -and $code -in 32, 33, 1175, 1176) { return 5 }
    if ($e -is [UnauthorizedAccessException] -or ($e -is [IO.IOException] -and $code -eq 5)) {
        if ($target -and (Test-Path -LiteralPath $target -PathType Leaf) -and ((Get-Item -LiteralPath $target -Force).Attributes -band [IO.FileAttributes]::ReadOnly)) { return 1 }
        return 3
    }
    return 1
}
# $ioOp does the work. Before a retry, $ioDone says whether an earlier attempt completed after all (then it
# is not repeated) and $ioGuard throws when retrying is no longer safe.
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
# Test hooks (-TestMoveFailures / -TestReplaceFailures [-TestReplaceTarget <file name>]): simulated transient failures.
function Test-IoHook([string]$hook, [string]$target) {
    if ($hook -eq 'move' -and $script:TestMoveFailures -gt 0) { $script:TestMoveFailures--; throw [IO.IOException]::new("simulated transient failure moving $target", -2147024864) }
    if ($hook -eq 'replace' -and $script:TestReplaceFailures -gt 0 -and (-not $TestReplaceTarget -or [IO.Path]::GetFileName($target) -ieq $TestReplaceTarget)) {
        $script:TestReplaceFailures--; throw [IO.IOException]::new("simulated: Unable to remove the file to be replaced ($target)", -2147023721)
    }
}

# Atomic replace: write a temporary file in the same folder, then swap it in.
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
$EnvPrefix = 'PPTK-INSTALL'
. (Join-Path $PSScriptRoot 'deploy-env.ps1')
if ($Mode -eq 'production' -and -not $PackagePath -and -not $Rollback) { $PackagePath = $PSScriptRoot }
$PptkDir = Join-Path $LocalAppData 'PPTK'
$AddinsDir = Join-Path $AppData 'Microsoft\AddIns'
$AddinFile = Join-Path $AddinsDir 'PPTK.ppam'
$StatePath = Join-Path $PptkDir 'state.json'
$JournalPath = Join-Path $PptkDir 'transaction.json'
$DetectScript = Join-Path $PSScriptRoot 'detect.ps1'
$script:LogPath = $null
$script:HelperBusy = $false

function Test-IsCurrentUser([string]$v) {
    if (-not $v) { return $false }
    if ($v -ieq $Identity.sid -or $v -ieq $Identity.name) { return $true }
    return ($v -notmatch '\\') -and ($v -ieq (($Identity.name -split '\\')[-1]))
}

function Step-Hook([string]$step) {
    if ($TestCrashAt -eq $step) { Log "TEST: simulated crash after step $step"; [Environment]::Exit(99) }
    if ($TestFailAt -eq $step) { throw "simulated failure after step $step" }
}

# ---------------------------------------------------------------- package validation (before any change)
function Read-Package([string]$pkg) {
    if (-not $pkg -or -not (Test-Path -LiteralPath (Join-Path $pkg 'version.json'))) { Stop-Install 1 'InvalidPackage' "version.json not found in '$pkg'" }
    $pkg = (Get-Item -LiteralPath (Resolve-Path -LiteralPath $pkg).Path).FullName.TrimEnd('\')   # long form, as Get-ChildItem returns
    try { $m = [IO.File]::ReadAllText((Join-Path $pkg 'version.json')) | ConvertFrom-Json } catch { Stop-Install 1 'InvalidPackage' 'version.json is not valid JSON' }
    $bad = @()
    if (-not (Has $m 'schemaVersion') -or $m.schemaVersion -ne 1) { $bad += 'schemaVersion must be 1' }
    if (-not (Has $m 'packageName') -or $m.packageName -ne 'PPTK') { $bad += "packageName must be 'PPTK'" }
    if (-not (Has $m 'packageVersion') -or -not (Test-SemVer $m.packageVersion)) { $bad += 'packageVersion missing or not MAJOR.MINOR.PATCH' }
    if (-not (Has $m 'files') -or @($m.files).Count -eq 0) { $bad += 'files missing' }
    if ($bad.Count) { Stop-Install 1 'InvalidPackage' ($bad -join '; ') }
    $listed = @()
    foreach ($f in @($m.files)) {
        if (-not (Has $f 'path') -or -not ($f.path -is [string]) -or $f.path -notmatch '^[A-Za-z0-9._-]+(/[A-Za-z0-9._-]+)*$' -or $f.path -match '(^|/)\.\.?(/|$)') { Stop-Install 1 'InvalidPackage' "unsafe or missing file path '$(if (Has $f 'path') { $f.path })'" }
        if (-not (Has $f 'sha256') -or $f.sha256 -cnotmatch '^[0-9a-f]{64}$') { Stop-Install 1 'InvalidPackage' "invalid sha256 for $($f.path)" }
        if ($f.path -match '^payload/runtime/' -and $f.path -notmatch $RuntimePath) { Stop-Install 1 'InvalidPackage' "unknown runtime folder: $($f.path)" }
        if ($f.path -match $BannedExtensions -and $f.path -notmatch $RuntimePath) { Stop-Install 1 'InvalidPackage' "executable or banned file type in the package: $($f.path)" }
        $p = Join-Path $pkg ($f.path.Replace('/', '\'))
        if (-not (Test-Path -LiteralPath $p)) { Stop-Install 1 'InvalidPackage' "payload file missing: $($f.path)" }
        if ((Get-Sha256 $p) -ne $f.sha256) { Stop-Install 1 'InvalidPackage' "SHA-256 mismatch: $($f.path)" }
        $listed += $f.path
    }
    foreach ($req in @('payload/PPTK.ppam') + @($KeysFiles | ForEach-Object { "payload/PPTKKeys/$_" })) { if ($listed -notcontains $req) { Stop-Install 1 'InvalidPackage' "required file not listed: $req" } }
    $onDisk = @(Get-ChildItem -LiteralPath $pkg -Recurse -File -Force | ForEach-Object { $_.FullName.Substring($pkg.Length + 1).Replace('\', '/') })
    $extra = @($onDisk | Where-Object { $_ -ne 'version.json' -and $listed -notcontains $_ })
    if ($extra.Count) { Stop-Install 1 'InvalidPackage' ("files not listed in version.json: " + ($extra -join ', ')) }
    # Executables only inside the official .NET runtimes (signature-checked when staged).
    $exe = @($onDisk | Where-Object { $_ -match $BannedExtensions -and $_ -notmatch $RuntimePath })
    if ($exe.Count) { Stop-Install 1 'InvalidPackage' ("executable or banned file type in the package: " + ($exe -join ', ')) }
    $hash = @{}; foreach ($f in @($m.files)) { $hash[$f.path] = $f.sha256 }
    try { $fw = ([IO.File]::ReadAllText((Join-Path $pkg 'payload\PPTKKeys\PPTKKeys.runtimeconfig.json')) | ConvertFrom-Json).runtimeOptions.framework } catch { Stop-Install 1 'InvalidPackage' 'PPTKKeys.runtimeconfig.json is not valid JSON' }
    if (-not $fw -or -not $fw.name -or -not (Test-SemVer ([string]$fw.version))) { Stop-Install 1 'InvalidPackage' 'PPTKKeys.runtimeconfig.json has no framework name/version' }
    # Universal package: a private runtime per architecture (payload/runtime/<rid>/, with a fingerprint in "runtimes").
    $rids = @($listed | ForEach-Object { if ($_ -match $RuntimePath) { $Matches[1] } } | Select-Object -Unique)
    foreach ($r in $rids) { if (-not ((Has $m 'runtimes') -and (Has $m.runtimes $r) -and $m.runtimes.$r.fingerprint -cmatch '^[0-9a-f]{64}$')) { Stop-Install 1 'InvalidPackage' "version.json has no runtime fingerprint for $r" } }
    return [pscustomobject]@{ Dir = $pkg; Manifest = $m; Version = $m.packageVersion; Hash = $hash; Framework = $fw.name; MinRuntime = [version]$fw.version; Runtimes = $rids; Rid = $null }
}

# The manifest entries of one architecture's runtime: @{ Rel = path under runtime\; Key = manifest path }.
function Get-RuntimeEntries($hash, [string]$rid) {
    $prefix = "payload/runtime/$rid/"
    return @($hash.Keys | Where-Object { $_.StartsWith($prefix, [StringComparison]::Ordinal) } | Sort-Object | ForEach-Object { [pscustomobject]@{ Rel = $_.Substring($prefix.Length).Replace('/', '\'); Key = $_ } })
}

# .NET runtime: dotnet.exe plus shared\<framework>\<version> with the same major version and at least
# the minimum, searched in the INSTALLER.md order (per-user, then machine-wide). Never executed.
function Find-Dotnet($pkg) {
    foreach ($dir in $DotnetDirs) {
        $exe = Join-Path $dir 'dotnet.exe'
        $shared = Join-Path $dir "shared\$($pkg.Framework)"
        if (-not (Test-Path -LiteralPath $exe) -or -not (Test-Path -LiteralPath $shared)) { continue }
        $ok = @(Get-ChildItem -LiteralPath $shared -Directory | Where-Object { $_.Name -match '^\d+\.\d+\.\d+$' } | Where-Object { $v = [version]$_.Name; $v.Major -eq $pkg.MinRuntime.Major -and $v -ge $pkg.MinRuntime })
        if ($ok.Count) { return $exe }
    }
    Stop-Install 1 'MissingRuntime' "$($pkg.Framework) $($pkg.MinRuntime) (same major version) was not found; install the .NET runtime first"
}

# ---------------------------------------------------------------- current installation
function Get-CurrentInstall {
    $tasks = Get-Tasks
    $main = @($tasks | Where-Object { $_.name -eq $TaskName -and $_.path -eq $TaskPath })
    $dups = @($tasks | Where-Object { -not ($_.name -eq $TaskName -and $_.path -eq $TaskPath) -and @(@($_.actions) | Where-Object { "$($_.execute) $($_.arguments)" -match 'PPTKKeys\.dll' }).Count })
    $stateRaw = $null; $state = $null; $stateCorrupt = $false
    if (Test-Path -LiteralPath $StatePath) { try { $state = [IO.File]::ReadAllText($StatePath) | ConvertFrom-Json; if (-not (Has $state 'installed')) { $stateCorrupt = $true } } catch { $stateCorrupt = $true } }
    $installed = Get-Reg $PptkKey 'InstalledVersion'
    $kind = 'Fresh'; $approval = @()
    if ($dups.Count) { $approval += ('other tasks already run PPTKKeys: ' + (($dups | ForEach-Object { "$($_.path)$($_.name)" }) -join ', ')) }
    if ($main.Count) {
        $t = $main[0]; $a = @($t.actions)
        if (-not (Test-IsCurrentUser $t.userId)) { $approval += "task '$TaskName' belongs to '$($t.userId)', not the current user" }
        if ($a.Count -ne 1 -or ([IO.Path]::GetFileName(([string]$a[0].execute).Trim('"'))) -ine 'dotnet.exe' -or [string]$a[0].arguments -notmatch '^\s*exec\s+"[^"]+PPTKKeys\.dll"\s*$') { $approval += "task '$TaskName' does not run dotnet.exe exec ""...PPTKKeys.dll""" }
    }
    if (Test-RegKey $AddinKey) {
        $rp = [string](Get-Reg $AddinKey 'Path')
        if (-not ($rp -ieq 'PPTK.ppam' -or (Same-Path $rp $AddinFile))) { $approval += "add-in registration points to '$rp'" }
    }
    if ($stateCorrupt) { $approval += 'state.json is unreadable or incomplete' }
    if ($installed -or $state) {
        $kind = 'Managed'
        if ($installed -and -not (Test-SemVer ([string]$installed))) { $approval += "InstalledVersion '$installed' is not a version" }
        if ($state -and -not $stateCorrupt -and $installed -and $state.installed -and $state.installed -ne $installed) { $approval += "state.json ($($state.installed)) and InstalledVersion ($installed) disagree" }
    } elseif ((Test-Path -LiteralPath $AddinFile) -or $main.Count -or (Test-RegKey $AddinKey) -or (Test-Path -LiteralPath (Join-Path $PptkDir 'PPTKKeys'))) {
        $kind = 'Legacy'
    }
    return [pscustomobject]@{ Kind = $kind; Installed = $(if ($state -and -not $stateCorrupt) { $state.installed } else { $installed }); State = $state; Task = $(if ($main.Count) { $main[0] } else { $null }); Approval = $approval }
}

# ---------------------------------------------------------------- journal (undo of the running operation)
function Save-Journal([string]$operation, [string]$target) {
    $txn = Join-Path $PptkDir 'txn'; New-Item -ItemType Directory -Force -Path $txn | Out-Null
    $addinBackup = $null
    if (Test-Path -LiteralPath $AddinFile) { $addinBackup = Join-Path $txn 'PPTK.ppam.before'; Copy-Item -LiteralPath $AddinFile -Destination $addinBackup -Force }
    $j = [ordered]@{
        journalVersion = 1; tasksFormat = $TasksFormat   # src\Deploy\JOURNAL.md
        operation = $operation; target = $target; startedUtc = [DateTime]::UtcNow.ToString('o')
        stateBefore = $(if (Test-Path -LiteralPath $StatePath) { [IO.File]::ReadAllText($StatePath) } else { $null })
        tasksBefore = (Export-Tasks)
        pptkKeyBefore = (Export-RegKey $PptkKey)
        addinKeyBefore = (Export-RegKey $AddinKey)
        addinBefore = $addinBackup; addinBeforeSha = $(if ($addinBackup) { Get-Sha256 $addinBackup } else { $null })
        createdVersionDir = $(if (Test-Path -LiteralPath (Join-Path $PptkDir "versions\$target")) { $null } else { "versions\$target" })
    }
    Write-Atomic $JournalPath (ConvertTo-JsonText $j)
    Log "journal written for $operation $target"
}

# Puts everything back exactly as recorded before the operation, then removes the journal.
function Undo-Journal([string]$why) {
    $j = Read-Json $JournalPath
    if (-not $j) { return }
    Log "undo ($why): restoring the state from before $($j.operation) $($j.target)"
    Assert-JournalFormat $j
    Stop-HelperForUndo
    # An interrupted uninstall.ps1 moved folders into txn\removed\: move them back first.
    foreach ($mv in @($j.moved | Where-Object { $_ })) {
        $src = Join-Path $PptkDir $mv.to; $dst = Join-Path $PptkDir $mv.from
        if ((Test-Path -LiteralPath $src) -and -not (Test-Path -LiteralPath $dst)) { Move-WithRetry $src $dst }
        if (-not (Test-Path -LiteralPath $dst)) { throw "$($mv.from) could not be restored" }
    }
    if ($j.addinBefore) {
        if ((Get-Sha256 $j.addinBefore) -ne $j.addinBeforeSha) { throw 'the add-in backup in the journal is damaged' }
        # Unchanged: do not write over a file PowerPoint may have opened since.
        if (-not (Test-Path -LiteralPath $AddinFile) -or (Get-Sha256 $AddinFile) -ne $j.addinBeforeSha) { Copy-Item -LiteralPath $j.addinBefore -Destination $AddinFile -Force }
        if ((Get-Sha256 $AddinFile) -ne $j.addinBeforeSha) { throw 'the add-in could not be restored' }
    } elseif (Test-Path -LiteralPath $AddinFile) { Remove-Item -LiteralPath $AddinFile -Force }
    Import-Tasks $j.tasksBefore
    if ($null -ne $j.pptkKeyBefore) { Set-RegKey $PptkKey ($j.pptkKeyBefore | ConvertFrom-Json) } else { Set-RegKey $PptkKey $null }
    if ($null -ne $j.addinKeyBefore) { Set-RegKey $AddinKey ($j.addinKeyBefore | ConvertFrom-Json) } else { Set-RegKey $AddinKey $null }
    if ($null -ne $j.stateBefore) { Write-Atomic $StatePath $j.stateBefore } elseif (Test-Path -LiteralPath $StatePath) { Remove-Item -LiteralPath $StatePath -Force }
    Remove-CreatedVersionDir $j.createdVersionDir
    Complete-Journal
    [void](Helper-Action 'restart (restored)')
    Log 'undo complete'
}
# The journal is deleted first: that is the commit (a journal left behind would be undone by the next run,
# so a failure here is an error). txn\ is best effort; a leftover one without a journal is removed later.
function Complete-Journal {
    if (Test-Path -LiteralPath $JournalPath) { Remove-Item -LiteralPath $JournalPath -Force }
    Remove-Item -LiteralPath (Join-Path $PptkDir 'txn') -Recurse -Force -ErrorAction SilentlyContinue
}

# Folder moves right after the files were written (same bounded retry; idempotent: a move that completed
# is not repeated).
function Move-WithRetry([string]$from, [string]$to) {
    Invoke-IoRetry "move $([IO.Path]::GetFileName($from))" $from 'move' `
        { Move-Item -LiteralPath $from -Destination $to -ErrorAction Stop } `
        { -not (Test-Path -LiteralPath $from) -and (Test-Path -LiteralPath $to) } `
        { if (-not (Test-Path -LiteralPath $from)) { throw "$from disappeared before it could be moved" } }
}

# ---------------------------------------------------------------- steps shared by install and rollback
# Files of versions\<v>: relative path -> package key. Universal package: only this device's runtime, as runtime\.
function Get-StageSources($pkg) {
    $sources = @{ 'PPTK.ppam' = 'payload/PPTK.ppam' }; foreach ($f in $KeysFiles) { $sources["PPTKKeys\$f"] = "payload/PPTKKeys/$f" }
    if ($pkg.Rid) { foreach ($e in (Get-RuntimeEntries $pkg.Hash $pkg.Rid)) { $sources["runtime\$($e.Rel)"] = $e.Key } }
    return $sources
}

# $true when versions\<v> exists but does not match the package, so Stage-Version will replace it.
function Test-StageReplaces($pkg) {
    $dest = Join-Path $PptkDir "versions\$($pkg.Version)"
    if (-not (Test-Path -LiteralPath $dest)) { return $false }
    $sources = Get-StageSources $pkg
    foreach ($k in $sources.Keys) { $p = Join-Path $dest $k; if (-not (Test-Path -LiteralPath $p) -or (Get-Sha256 $p) -ne $pkg.Hash[$sources[$k]]) { return $true } }
    return $false
}

function Stage-Version($pkg) {
    $dest = Join-Path $PptkDir "versions\$($pkg.Version)"
    $sources = Get-StageSources $pkg
    if (Test-Path -LiteralPath $dest) {
        if (-not (Test-StageReplaces $pkg)) { Log "versions\$($pkg.Version) already staged and intact"; return $dest }
        Remove-Item -LiteralPath $dest -Recurse -Force
    }
    $tmp = "$dest.staging-" + [guid]::NewGuid().ToString('N').Substring(0, 6)
    New-Item -ItemType Directory -Force -Path (Join-Path $tmp 'PPTKKeys') | Out-Null
    foreach ($k in $sources.Keys) { $to = Join-Path $tmp $k; $toDir = Split-Path -Parent $to; if (-not (Test-Path -LiteralPath $toDir)) { New-Item -ItemType Directory -Force -Path $toDir | Out-Null }; Copy-Item -LiteralPath (Join-Path $pkg.Dir ($sources[$k].Replace('/', '\'))) -Destination $to }
    Copy-Item -LiteralPath (Join-Path $pkg.Dir 'version.json') -Destination (Join-Path $tmp 'version.json')
    foreach ($k in $sources.Keys) { if ((Get-Sha256 (Join-Path $tmp $k)) -ne $pkg.Hash[$sources[$k]]) { throw "staged copy of $k does not match the package" } }
    if ($pkg.Rid) {
        $rt = Join-Path $tmp 'runtime'
        try { Assert-Runtime $rt $pkg.Rid $pkg.Framework $pkg.MinRuntime } catch { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue; throw "InvalidRuntime: $($_.Exception.Message)" }   # a failure: the journal is undone
        if ((Get-RuntimeFingerprint $rt) -ne $pkg.Manifest.runtimes.($pkg.Rid).fingerprint) { throw "the staged $($pkg.Rid) runtime does not match its fingerprint in version.json" }
        Log "private .NET runtime $($pkg.Rid) staged and verified (Microsoft signatures, architecture, fingerprint)"
    }
    Move-WithRetry $tmp $dest
    Log "staged versions\$($pkg.Version)"
    return $dest
}

# Point the one PPTKKeys task at a version folder; the existing task is updated in place (adopted).
# The helper is stopped and started only when its files change ($restart); otherwise the running
# helper keeps its (kept) previous folder and the new path is used from the next logon.
function Set-HelperTask([string]$dotnet, [string]$dll, [bool]$restart = $true) {
    $current = @(Get-Tasks | Where-Object { $_.name -eq $TaskName -and $_.path -eq $TaskPath })
    $want = New-TaskEntry $dotnet $dll
    $same = $current.Count -and ([string]@($current[0].actions)[0].execute -eq $dotnet) -and ([string]@($current[0].actions)[0].arguments -eq $want.actions[0].arguments) -and (Test-IsCurrentUser $current[0].userId) -and $current[0].enabled -and $current[0].logonType -eq 'Interactive' -and $current[0].runLevel -eq 'Limited'
    if ($same) { Log 'task already correct'; return $false }
    # The running helper is asked to close (never terminated). If it does not, the task still points at the
    # new helper, which then starts at the next sign-in; the old helper's folder is kept until then.
    if ($restart -and $current.Count) { if (-not (Helper-Action 'stop')) { $script:HelperBusy = $true; Log 'the running helper did not close; the new helper starts at the next sign-in' } }
    Set-MainTask $dotnet $dll
    if ($restart) { [void](Helper-Action 'start') }
    Log "task now runs $dll$(if (-not $restart) { ' (helper files unchanged: no restart)' })"
    return $true
}

function Set-Registration { Set-RegKey $AddinKey ([ordered]@{ Path = 'PPTK.ppam'; AutoLoad = 1 }) }

# $true when the add-in already has this hash, or can be replaced now (PowerPoint cannot be holding it).
function Test-AddinReplaceable([string]$sha) {
    if (-not (Test-Path -LiteralPath $AddinFile)) { return $true }
    if ((Get-Sha256 $AddinFile) -eq $sha) { return $true }
    if (Test-PowerPointRunning) { return $false }
    try { $s = [IO.File]::Open($AddinFile, 'Open', 'Read', 'None'); $s.Dispose() } catch { return $false }   # read access: a read-only add-in is not "in use"
    return $true
}

# Replace the add-in only when PowerPoint cannot be holding it. Returns $true when replaced.
function Install-Addin([string]$source, [string]$sha) {
    if ((Test-Path -LiteralPath $AddinFile) -and (Get-Sha256 $AddinFile) -eq $sha) { return $true }
    if (-not (Test-AddinReplaceable $sha)) { return $false }
    New-Item -ItemType Directory -Force -Path $AddinsDir | Out-Null
    $tmp = Join-Path $AddinsDir ('PPTK.ppam.new-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
    Copy-Item -LiteralPath $source -Destination $tmp
    if ((Get-Sha256 $tmp) -ne $sha) { Remove-Item -LiteralPath $tmp -Force; throw 'the add-in copy does not match the package' }
    # Retried only while PowerPoint is still not running: the add-in is never replaced under PowerPoint.
    Invoke-IoRetry 'replace PPTK.ppam' $AddinFile 'replace' `
        { if (Test-Path -LiteralPath $AddinFile) { [IO.File]::Replace($tmp, $AddinFile, [NullString]::Value) } else { [IO.File]::Move($tmp, $AddinFile) } } `
        { -not (Test-Path -LiteralPath $tmp) -and (Test-Path -LiteralPath $AddinFile) -and (Get-Sha256 $AddinFile) -eq $sha } `
        {
            if (Test-PowerPointRunning) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue; throw 'PowerPoint started while the add-in was being replaced' }
            if (-not (Test-Path -LiteralPath $tmp)) { throw 'the add-in copy disappeared before it could be installed' }
        }
    if ((Get-Sha256 $AddinFile) -ne $sha) { throw 'the installed add-in does not match the package after replacing it' }
    return $true
}

function Write-State($obj) { Write-Atomic $StatePath (ConvertTo-JsonText $obj) }

function Invoke-Detect([string]$manifest) {
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$DetectScript`" -ManifestPath `"$manifest`"" + $(if ($Root) { " -Root `"$Root`"" } else { '' }) + $(if ($LiveTestId) { " -LiveTestId $LiveTestId" } else { '' })
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8; $psi.StandardErrorEncoding = [Text.Encoding]::UTF8
    $p = [Diagnostics.Process]::Start($psi); $o = $p.StandardOutput.ReadToEnd(); $e = $p.StandardError.ReadToEnd(); $p.WaitForExit()
    return [pscustomobject]@{ Code = $p.ExitCode; Out = $o.Trim(); Err = $e.Trim() }
}

function Remove-OldVersions([string[]]$keep) {
    $vd = Join-Path $PptkDir 'versions'
    if (-not (Test-Path -LiteralPath $vd)) { return }
    foreach ($d in @(Get-ChildItem -LiteralPath $vd -Directory)) {
        if ($keep -notcontains $d.Name) {
            # Best effort, after the commit: a folder a running helper still uses is removed by a later run.
            try { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction Stop; Log "removed old versions\$($d.Name)" }
            catch { Log "old versions\$($d.Name) kept for now: $($_.Exception.Message)" }
        }
    }
}

# ---------------------------------------------------------------- operations
function Invoke-Install {
    $pkg = Read-Package $PackagePath
    if ($pkg.Runtimes.Count) {
        # Universal package: the private runtime of this device's architecture; no installed .NET is needed.
        $pkg.Rid = Get-NativeRid
        if ($pkg.Runtimes -notcontains $pkg.Rid) { Stop-Install 1 'UnsupportedArchitecture' "this device is $($pkg.Rid); the package has runtimes for $($pkg.Runtimes -join ', ')" }
        $dotnet = $null
        Log "package PPTK $($pkg.Version) valid (universal: $($pkg.Runtimes -join ', ')); this device: $($pkg.Rid)"
    } else {
        $dotnet = Find-Dotnet $pkg
        Log "package PPTK $($pkg.Version) valid; .NET host $dotnet"
    }
    $cur = Get-CurrentInstall
    Log "current installation: $($cur.Kind) $($cur.Installed)"
    if ($cur.Approval.Count) { Stop-Install 1 'ManualApprovalRequired' ($cur.Approval -join '; ') }
    if ($cur.Kind -eq 'Managed' -and $cur.Installed -and [version]$cur.Installed -gt [version]$pkg.Version) {
        Stop-Install 1 'Downgrade' "installed $($cur.Installed) is newer than the package $($pkg.Version); use -Rollback to return to the previous version"
    }
    # Same version already complete (as detect.ps1 sees it): change nothing at all.
    if ($cur.Kind -eq 'Managed' -and $cur.Installed -eq $pkg.Version) {
        $already = Invoke-Detect (Join-Path $pkg.Dir 'version.json')
        if ($already.Code -eq 0) { Log 'already installed and healthy; nothing to do'; return "PPTK $($pkg.Version) already installed (no change)" }
        Log "installed $($pkg.Version) needs repair: $($already.Err)"
    }

    # Same-version repair that replaces versions\<v> while the task runs the helper from it: the running helper
    # locks its files, so it is asked to close first (before the journal: a helper that does not close changes
    # nothing) and is started again once the folder is replaced.
    $helperStopped = $false
    if ($cur.Task -and (Test-StageReplaces $pkg)) {
        $vdirPath = Join-Path $PptkDir "versions\$($pkg.Version)"
        $a0 = @($cur.Task.actions)[0]
        $usesDir = (([string]$a0.execute) + ' ' + ([string]$a0.arguments)).IndexOf($vdirPath.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase) -ge 0
        if ($usesDir) {
            if (-not (Helper-Action 'stop')) { Stop-Install 1618 'HelperBusy' "the running helper uses versions\$($pkg.Version) and did not close; nothing was changed, retry later" }
            $helperStopped = $true
            Log "the helper runs from versions\$($pkg.Version), which is replaced: closed for the repair"
        }
    }

    Save-Journal 'install' $pkg.Version
    $vdir = Stage-Version $pkg; Step-Hook 'staged'
    $dll = Join-Path $vdir 'PPTKKeys\PPTKKeys.dll'
    if ($pkg.Rid) { $dotnet = Join-Path $vdir 'runtime\dotnet.exe' }
    # Helper files or runtime changed? Compare with what the current task runs (the PPTKKeys engine checks the same).
    $keysChanged = $true
    if ($cur.Task) {
        $m0 = [regex]::Match([string]@($cur.Task.actions)[0].arguments, '"([^"]+)"')
        if ($m0.Success) {
            $oldDir = Split-Path -Parent $m0.Groups[1].Value
            $keysChanged = $false
            foreach ($f in $KeysFiles) { $p = Join-Path $oldDir $f; if (-not (Test-Path -LiteralPath $p) -or (Get-Sha256 $p) -ne $pkg.Hash["payload/PPTKKeys/$f"]) { $keysChanged = $true } }
            if ($pkg.Rid -and -not (Same-Path ([string]@($cur.Task.actions)[0].execute).Trim('"') $dotnet)) { $keysChanged = $true }
        }
    }
    $arch = if ($pkg.Rid) { $pkg.Rid } elseif ($cur.State -and (Has $cur.State 'architecture')) { $cur.State.architecture } else { $null }
    # The task, registration and add-in switch together, or not at all: while PowerPoint holds the
    # old add-in, the installed version stays complete (its task and helper are not touched) and the
    # new version waits in versions\ + pending\ until a run with PowerPoint closed applies all of it.
    $applied = Test-AddinReplaceable $pkg.Hash['payload/PPTK.ppam']
    $helperStarted = $false
    if ($applied) {
        $helperStarted = (Set-HelperTask $dotnet $dll $keysChanged) -and $keysChanged; Step-Hook 'task'
        Set-Registration; Step-Hook 'registration'
        if (-not (Install-Addin (Join-Path $vdir 'PPTK.ppam') $pkg.Hash['payload/PPTK.ppam'])) { throw 'PowerPoint opened the add-in during the update' }
        Step-Hook 'addin'
    }
    # The helper closed for the repair runs again from the replaced folder (also while the add-in update is pending).
    if ($helperStopped -and -not $helperStarted) { [void](Helper-Action 'start') }

    $previous = if ($cur.Kind -eq 'Managed' -and $cur.Installed -and $cur.Installed -ne $pkg.Version) { $cur.Installed } elseif ($cur.State -and (Has $cur.State 'previous')) { $cur.State.previous } else { $null }
    $legacy = if ($cur.State -and (Has $cur.State 'legacyBackup')) { $cur.State.legacyBackup } else { $null }
    if ($cur.Kind -eq 'Legacy') {
        $lb = Join-Path $PptkDir ('legacy\' + [DateTime]::UtcNow.ToString('yyyyMMddHHmmss'))
        New-Item -ItemType Directory -Force -Path $lb | Out-Null
        $jb = (Read-Json $JournalPath).addinBefore
        if ($jb) { Copy-Item -LiteralPath $jb -Destination (Join-Path $lb 'PPTK.ppam') }
        if ($cur.Task) { [IO.File]::WriteAllText((Join-Path $lb 'task.json'), (ConvertTo-JsonText $cur.Task), $Utf8) }
        $legacy = $lb.Substring($PptkDir.Length + 1)
        Log "legacy installation adopted; its add-in and task are kept in $legacy, the old PPTKKeys folder is left in place"
    }

    if ($applied) {
        Write-State ([ordered]@{ schemaVersion = 1; installed = $pkg.Version; previous = $previous; pending = $null; legacyBackup = $legacy; architecture = $arch; lastError = $null; updatedUtc = [DateTime]::UtcNow.ToString('o') })
        Set-RegKey $PptkKey ([ordered]@{ InstalledVersion = $pkg.Version })
        Remove-Item -LiteralPath (Join-Path $PptkDir 'pending') -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        # The same staged version keeps its "notified" flag, so the employee is told only once.
        $wasNotified = ($cur.State -and (Has $cur.State 'pending') -and $cur.State.pending -and $cur.State.pending.version -eq $pkg.Version -and $cur.State.pending.notified)
        $pend = Join-Path $PptkDir 'pending'; New-Item -ItemType Directory -Force -Path $pend | Out-Null
        $pfile = Join-Path $pend "PPTK-$($pkg.Version).ppam"
        Copy-Item -LiteralPath (Join-Path $vdir 'PPTK.ppam') -Destination $pfile -Force
        if ((Get-Sha256 $pfile) -ne $pkg.Hash['payload/PPTK.ppam']) { throw 'the staged pending add-in does not match the package' }
        Get-ChildItem -LiteralPath $pend -File | Where-Object { $_.Name -ne "PPTK-$($pkg.Version).ppam" } | Remove-Item -Force
        $keepInstalled = if ($cur.Installed) { $cur.Installed } else { $null }
        $keepRolledBack = if ($cur.State -and (Has $cur.State 'rolledBackFrom')) { $cur.State.rolledBackFrom } else { $null }
        Write-State ([ordered]@{ schemaVersion = 1; installed = $keepInstalled; previous = $(if ($cur.State -and (Has $cur.State 'previous')) { $cur.State.previous } else { $null }); rolledBackFrom = $keepRolledBack; pending = [ordered]@{ version = $pkg.Version; file = "pending\PPTK-$($pkg.Version).ppam"; sha256 = $pkg.Hash['payload/PPTK.ppam']; keysChanged = [bool]$keysChanged; notified = [bool]$wasNotified; stagedUtc = [DateTime]::UtcNow.ToString('o') }; legacyBackup = $legacy; architecture = $arch; lastError = $null; updatedUtc = [DateTime]::UtcNow.ToString('o') })
        Log "PowerPoint has the add-in open: PPTK $($pkg.Version) staged in versions\ and pending\; the installed version, its task and the helper are unchanged"
    }
    Step-Hook 'state'

    $det = Invoke-Detect (Join-Path $vdir 'version.json')
    if ($applied -and $det.Code -ne 0) { throw "detect.ps1 does not confirm the installation: $($det.Err)" }
    if (-not $applied -and ($det.Code -eq 0 -or $det.Err -notmatch 'PendingUpdate')) { throw "detect.ps1 does not report the pending update: $($det.Out) $($det.Err)" }
    if (-not $applied -and $cur.Kind -eq 'Managed' -and $cur.Installed) {
        $oldManifest = Join-Path $PptkDir "versions\$($cur.Installed)\version.json"
        if (Test-Path -LiteralPath $oldManifest) {
            $old = Invoke-Detect $oldManifest
            if ($old.Code -ne 0) { throw "detect.ps1 no longer confirms the installed $($cur.Installed) while the update is pending: $($old.Err)" }
        }
    }
    Complete-Journal
    # Pending: the installed version and ITS previous stay usable (rollback still goes one step back).
    $keepPrev = if (-not $applied -and $cur.State -and (Has $cur.State 'previous')) { $cur.State.previous } else { $null }
    $keep = @($pkg.Version) + @($previous, $(if ($cur.Installed) { $cur.Installed }), $keepPrev) | Where-Object { $_ } | Select-Object -Unique
    Remove-OldVersions $keep
    if (-not $applied) { Stop-Install 1618 'Pending' "PPTK $($pkg.Version) is staged; it is applied when PowerPoint is closed (installed version stays $($cur.Installed))" }
    return "PPTK $($pkg.Version) installed ($($cur.Kind.ToLower()) $(if ($cur.Installed) { $cur.Installed }))" + $(if ($script:HelperBusy) { "; the running helper did not close, the new helper starts at the next sign-in" } else { "" })
}

function Invoke-Rollback {
    $cur = Get-CurrentInstall
    if ($cur.Kind -ne 'Managed' -or -not $cur.State) { Stop-Install 1 'RollbackUnavailable' 'no managed installation with a recorded state' }
    if ($cur.Approval.Count) { Stop-Install 1 'ManualApprovalRequired' ($cur.Approval -join '; ') }
    # An update that is only staged is cancelled; the installed version is kept (never two steps back).
    if ((Has $cur.State 'pending') -and $cur.State.pending) { return (Invoke-CancelPending $cur) }
    $installed = $cur.State.installed
    $prev = if (Has $cur.State 'previous') { $cur.State.previous } else { $null }
    $from = if ((Has $cur.State 'rolledBackFrom') -and $cur.State.rolledBackFrom) { " (already rolled back from $($cur.State.rolledBackFrom))" } else { '' }
    if (-not $prev) { Stop-Install 1 'RollbackUnavailable' "state.json records no previous version$from" }
    if (-not (Test-SemVer $prev) -or -not (Test-SemVer $installed) -or [version]$prev -ge [version]$installed) { Stop-Install 1 'RollbackUnavailable' "previous version '$prev' is not older than the installed '$installed'" }
    $vdir = Join-Path $PptkDir "versions\$prev"
    $manifest = Join-Path $vdir 'version.json'
    if (-not (Test-Path -LiteralPath $manifest)) { Stop-Install 1 'RollbackUnavailable' "versions\$prev is missing" }
    try { $m = [IO.File]::ReadAllText($manifest) | ConvertFrom-Json } catch { Stop-Install 1 'RollbackCorrupt' "versions\$prev\version.json is not valid JSON" }
    $hash = @{}; foreach ($f in @($m.files)) { $hash[$f.path] = $f.sha256 }
    $check = @{ 'PPTK.ppam' = 'payload/PPTK.ppam' }; foreach ($f in $KeysFiles) { $check["PPTKKeys\$f"] = "payload/PPTKKeys/$f" }
    foreach ($k in $check.Keys) {
        $p = Join-Path $vdir $k
        if (-not (Test-Path -LiteralPath $p)) { Stop-Install 1 'RollbackCorrupt' "versions\$prev\$k is missing" }
        if ((Get-Sha256 $p) -ne $hash[$check[$k]]) { Stop-Install 1 'RollbackCorrupt' "versions\$prev\$k does not match its manifest" }
    }
    $pkgLike = [pscustomobject]@{ Framework = ([IO.File]::ReadAllText((Join-Path $vdir 'PPTKKeys\PPTKKeys.runtimeconfig.json')) | ConvertFrom-Json).runtimeOptions.framework.name; MinRuntime = [version](([IO.File]::ReadAllText((Join-Path $vdir 'PPTKKeys\PPTKKeys.runtimeconfig.json')) | ConvertFrom-Json).runtimeOptions.framework.version) }
    # A universal version runs on its own private runtime (verified again against its manifest).
    $own = Join-Path $vdir 'runtime\dotnet.exe'
    $prevArch = if (Has $cur.State 'architecture') { $cur.State.architecture } else { $null }
    if (Test-Path -LiteralPath $own) {
        if (-not $RuntimeRids.ContainsKey([string]$prevArch)) { Stop-Install 1 'RollbackCorrupt' "state.json records no valid architecture for the runtime of versions\$prev" }
        foreach ($e in (Get-RuntimeEntries $hash $prevArch)) { $p = Join-Path $vdir "runtime\$($e.Rel)"; if (-not (Test-Path -LiteralPath $p) -or (Get-Sha256 $p) -ne $hash[$e.Key]) { Stop-Install 1 'RollbackCorrupt' "versions\$prev\runtime\$($e.Rel) is missing or does not match its manifest" } }
        $dotnet = $own
    } else { $dotnet = Find-Dotnet $pkgLike }
    if ((Test-Path -LiteralPath $AddinFile) -and (Get-Sha256 $AddinFile) -ne $hash['payload/PPTK.ppam'] -and (Test-PowerPointRunning)) { Stop-Install 1618 'RetryRollback' 'PowerPoint has the add-in open; close PowerPoint and run the rollback again (nothing was changed)' }

    Save-Journal 'rollback' $prev
    [void](Set-HelperTask $dotnet (Join-Path $vdir 'PPTKKeys\PPTKKeys.dll')); Step-Hook 'task'
    Set-Registration; Step-Hook 'registration'
    if (-not (Install-Addin (Join-Path $vdir 'PPTK.ppam') $hash['payload/PPTK.ppam'])) { throw 'the add-in is in use' }
    Step-Hook 'addin'
    # previous is cleared: a second rollback is refused instead of returning to the version just left.
    Write-State ([ordered]@{ schemaVersion = 1; installed = $prev; previous = $null; rolledBackFrom = $installed; pending = $null; legacyBackup = $(if (Has $cur.State 'legacyBackup') { $cur.State.legacyBackup }); architecture = $prevArch; lastError = $null; updatedUtc = [DateTime]::UtcNow.ToString('o') })
    Set-RegKey $PptkKey ([ordered]@{ InstalledVersion = $prev })
    Step-Hook 'state'
    $det = Invoke-Detect $manifest
    if ($det.Code -ne 0) { throw "detect.ps1 does not confirm the restored version: $($det.Err)" }
    Complete-Journal
    return "rolled back to PPTK $prev from $installed (detected)"
}

# Rollback while an update is staged: drop the staged update; task, helper, add-in and registration
# already belong to the installed version and are not touched.
function Invoke-CancelPending($cur) {
    $installed = $cur.State.installed
    $pv = $cur.State.pending.version
    if (-not (Test-SemVer $installed)) { Stop-Install 1 'RollbackUnavailable' "no installed version to keep; the pending update $pv was not cancelled" }
    $manifest = Join-Path $PptkDir "versions\$installed\version.json"
    if (-not (Test-Path -LiteralPath $manifest)) { Stop-Install 1 'RollbackUnavailable' "versions\$installed is missing; the pending update $pv was not cancelled" }
    Save-Journal 'cancel-pending' $installed
    $st = $cur.State
    $st.pending = $null
    $st | Add-Member -NotePropertyName lastError -NotePropertyValue $null -Force
    $st | Add-Member -NotePropertyName updatedUtc -NotePropertyValue ([DateTime]::UtcNow.ToString('o')) -Force
    Write-State $st; Step-Hook 'state'
    $det = Invoke-Detect $manifest
    if ($det.Code -ne 0) { throw "detect.ps1 does not confirm the installed $installed after cancelling: $($det.Err)" }
    Complete-Journal
    Remove-Item -LiteralPath (Join-Path $PptkDir 'pending') -Recurse -Force -ErrorAction SilentlyContinue
    Remove-OldVersions (@($installed, $(if (Has $st 'previous') { $st.previous })) | Where-Object { $_ })
    return "pending update $pv cancelled; PPTK $installed stays installed (detected)"
}

# ---------------------------------------------------------------- main
$lock = $null; $code = 0
try {
    New-Item -ItemType Directory -Force -Path (Join-Path $PptkDir 'logs') | Out-Null
    $script:LogPath = Join-Path $PptkDir ('logs\install-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + $PID + '.log')
    Log ("start: " + $(if ($Rollback) { 'rollback' } else { "install from $PackagePath" }))

    # state.lock: an exclusive handle, released by Windows even if this process is killed.
    $deadline = [DateTime]::UtcNow.AddSeconds($LockTimeoutSeconds)
    while (-not $lock) {
        try { $lock = [IO.File]::Open((Join-Path $PptkDir 'state.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch { if ([DateTime]::UtcNow -gt $deadline) { Stop-Install 1618 'Busy' 'another PPTK operation holds state.lock; nothing was changed' }; Start-Sleep -Milliseconds 250 }
    }

    if (Test-Path -LiteralPath $JournalPath) { Undo-Journal 'an earlier operation was interrupted' }
    else { Remove-Item -LiteralPath (Join-Path $PptkDir 'txn') -Recurse -Force -ErrorAction SilentlyContinue }   # left by a committed run

    $msg = if ($Rollback) { Invoke-Rollback } else { Invoke-Install }
    Log "success: $msg"
    [Console]::Out.WriteLine("PPTK-INSTALL OK: $msg")
} catch {
    $ex = $_.Exception; if ($ex -is [Management.Automation.MethodInvocationException] -and $ex.InnerException) { $ex = $ex.InnerException }
    if ($ex -is [InstallStop]) { $code = $ex.Code; $text = $ex.Message }
    else {
        $code = 1; $text = "PPTK-INSTALL Failed: $($ex.Message)"
        try { if (Test-Path -LiteralPath $JournalPath) { Undo-Journal 'the operation failed' } } catch { $text += " | RESTORE FAILED: $($_.Exception.Message) (transaction.json kept for the next run)" }
        try { $st = Read-Json $StatePath; if ($st) { $st | Add-Member -NotePropertyName lastError -NotePropertyValue $ex.Message -Force; Write-State $st } } catch { }
    }
    Log $text
    [Console]::Error.WriteLine($text)
} finally {
    if ($lock) { $lock.Dispose() }
}
exit $code
