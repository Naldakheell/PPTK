<#
.SYNOPSIS
  PPTK installation detection (Intune Win32 custom detection script). Strictly read-only.

.DESCRIPTION
  Intune contract (Microsoft Learn, "Add a Win32 app"): the app is detected when the script exits
  with 0 AND writes a string to STDOUT; any non-zero exit, or output on STDERR, means not detected.
  This script prints exactly one line to STDOUT only when the expected version is fully installed
  and healthy. Otherwise it exits 1, prints nothing to STDOUT, and writes one diagnostic line per
  problem to STDERR ("PPTK-DETECT <Category>: <detail>"); the first line is the primary reason.

  Intune runs a detection script with no parameters and without the package content, so the
  expected manifest (version.json: version and SHA-256 of every deployed file) is embedded by
  src\Deploy\build-package.ps1 into the copy it writes next to the package. This source file has
  no embedded manifest; pass -ManifestPath, or keep version.json next to the script.

  Checks for the current user (per-user installation):
    A  HKCU\Software\PPTK\InstalledVersion equals the manifest version
    B  %APPDATA%\Microsoft\AddIns\PPTK.ppam exists and its SHA-256 matches the manifest
    C  HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK: Path is "PPTK.ppam" or the full
       path of B, AutoLoad = 1
    D  %LOCALAPPDATA%\PPTK\versions\<version>\PPTKKeys\{PPTKKeys.dll, .deps.json, .runtimeconfig.json}
       exist and match the manifest
    E  exactly one task runs PPTKKeys: "\PPTKKeys Startup", enabled, principal = current user,
       interactive, limited, a logon trigger for the current user, action
       "<existing ...\dotnet.exe>" exec "<path of D's DLL>"
    F  %LOCALAPPDATA%\PPTK\state.json is valid, its "installed" equals the version, and there is no
       pending update for that version (a staged add-in is not an installation)
    G  universal package (manifest "runtimes"): the architecture recorded in state.json is this device's,
       versions\<version>\runtime\ matches that runtime's fingerprint, and the task (E) runs its dotnet.exe

  Categories: NotInstalled, OlderVersion, WrongVersion, PendingUpdate, MissingFiles, HashMismatch,
  MissingRegistration, IncorrectRegistration, IncorrectTask, IncorrectRuntime, CorruptState, InvalidManifest.

  Read-only: registry and task information is only read, files are opened for reading with sharing
  (PPTK.ppam can be open in PowerPoint), nothing is written, started, stopped or executed.

.PARAMETER Root
  Test mode. The environment is read from <Root> instead of the live user:
    <Root>\identity.json      {"name": "DOMAIN\\user", "sid": "S-..."}
    <Root>\registry.json      {"Software\\PPTK": {"InstalledVersion": "1.0.0"}, "<HKCU subkey>": {...}}
    <Root>\tasks.json         [ {name, path, enabled, userId, logonType, runLevel,
                                  triggers: [{type: "Logon", userId, enabled}],
                                  actions: [{execute, arguments}]} ]
    <Root>\AppData\...        stands for %APPDATA%
    <Root>\LocalAppData\...   stands for %LOCALAPPDATA%
  The same detection logic runs on it; the live registry, tasks and profile are never read.

.PARAMETER LiveTestId
  With -Root: live test mode (install.ps1 -LiveTestId). Folders from <Root>, but the real Task
  Scheduler and HKCU under test names: task "\PPTK-Test-<id>" (the only task considered), keys under
  HKCU\Software\PPTK-Test-<id>\. Read-only like every mode.
#>
[CmdletBinding()]
param(
    [string]$ManifestPath = '',
    [string]$Root = '',
    [string]$LiveTestId = ''
)

$ErrorActionPreference = 'Stop'

# Filled by build-package.ps1 in the generated detection script (Base64 of version.json).
$EmbeddedManifestBase64 = ''

$TaskName = 'PPTKKeys Startup'
$TaskPath = '\'
$RegPrefix = ''
if ($LiveTestId) {
    if (-not $Root -or $LiveTestId -cnotmatch '^[0-9a-f]{12}$') { [Console]::Error.WriteLine('PPTK-DETECT InvalidManifest: -LiveTestId needs -Root and 12 lowercase hex characters'); exit 1 }
    $TaskName = "PPTK-Test-$LiveTestId"
    $RegPrefix = "Software\PPTK-Test-$LiveTestId\"
}
$AddinKey = 'Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK'
$PptkKey = 'Software\PPTK'
$KeysFiles = @('PPTKKeys.dll', 'PPTKKeys.deps.json', 'PPTKKeys.runtimeconfig.json')

$script:Problems = New-Object System.Collections.ArrayList
function Problem([string]$category, [string]$detail) { [void]$script:Problems.Add("PPTK-DETECT ${category}: $detail") }

function Has([object]$obj, [string]$name) { return ($null -ne $obj) -and ($obj.PSObject.Properties.Name -contains $name) }
function Test-SemVer([object]$v) { return ($v -is [string]) -and ($v -match '^\d+\.\d+\.\d+$') }
function Test-Sha256([object]$v) { return ($v -is [string]) -and ($v -cmatch '^[0-9a-f]{64}$') }

# SHA-256 with read sharing, so an add-in that PowerPoint has open can still be checked.
function Get-FileSha256([string]$path) {
    $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '')).ToLowerInvariant() } finally { $sha.Dispose(); $stream.Dispose() }
}

function Same-Path([string]$a, [string]$b) {
    if (-not $a -or -not $b) { return $false }
    try { return [string]::Equals([IO.Path]::GetFullPath($a).TrimEnd('\'), [IO.Path]::GetFullPath($b).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) } catch { return $false }
}

# The machine's own architecture (not this process's: Intune may run this emulated on ARM64). Test roots may set
# it in <Root>\architecture.txt.
function Get-NativeRid {
    if ($Root) { $p = Join-Path $Root 'architecture.txt'; if (Test-Path -LiteralPath $p) { return ([IO.File]::ReadAllText($p)).Trim() } }
    $a = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | ForEach-Object { $_.Architecture } | Select-Object -First 1)
    if ($a.Count) { switch ([int]$a[0]) { 9 { return 'win-x64' } 12 { return 'win-arm64' } } }
    $n = (Get-ItemProperty -LiteralPath 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment' -Name PROCESSOR_ARCHITECTURE -ErrorAction SilentlyContinue).PROCESSOR_ARCHITECTURE
    switch ($n) { 'AMD64' { return 'win-x64' } 'ARM64' { return 'win-arm64' } }
    return 'unknown'
}

# Same definition as deploy-env.ps1 / build-package.ps1: SHA-256 over sorted "<relative path> <sha256>" lines.
function Get-RuntimeFingerprint([string]$dir) {
    $dir = (Get-Item -LiteralPath $dir).FullName.TrimEnd('\')   # the long form: Get-ChildItem returns long names under an 8.3 path
    $lines = [Collections.Generic.List[string]]::new()
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Recurse -File -Force)) { $lines.Add($f.FullName.Substring($dir.Length + 1).Replace('\', '/') + ' ' + (Get-FileSha256 $f.FullName)) }
    $lines.Sort([StringComparer]::Ordinal)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($lines -join "`n") + "`n"))).Replace('-', '')).ToLowerInvariant() } finally { $sha.Dispose() }
}

# ---------------------------------------------------------------- environment (live or test)
function Get-LiveEnvironment {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $tasks = @()
    foreach ($t in @(Get-ScheduledTask -ErrorAction SilentlyContinue)) {
        $actions = @($t.Actions | Where-Object { $_.CimClass.CimClassName -eq 'MSFT_TaskExecAction' } | ForEach-Object { [pscustomobject]@{ execute = $_.Execute; arguments = $_.Arguments } })
        $refers = ($t.TaskName -eq $TaskName) -or (-not $LiveTestId -and @($actions | Where-Object { "$($_.execute) $($_.arguments)" -match 'PPTKKeys\.dll' }).Count -gt 0)
        if (-not $refers) { continue }
        $triggers = @($t.Triggers | ForEach-Object {
            [pscustomobject]@{ type = $(if ($_.CimClass.CimClassName -eq 'MSFT_TaskLogonTrigger') { 'Logon' } else { $_.CimClass.CimClassName }); userId = [string]$_.UserId; enabled = [bool]$_.Enabled }
        })
        $tasks += [pscustomobject]@{
            name = $t.TaskName; path = $t.TaskPath; enabled = ($t.State -ne 'Disabled')
            userId = [string]$t.Principal.UserId; logonType = [string]$t.Principal.LogonType; runLevel = [string]$t.Principal.RunLevel
            triggers = $triggers; actions = $actions
        }
    }
    return [pscustomobject]@{
        Mode = 'live'; UserName = $id.Name; UserSid = $id.User.Value
        AppData = $(if ($LiveTestId) { Join-Path $Root 'AppData' } else { [Environment]::GetFolderPath('ApplicationData') })
        LocalAppData = $(if ($LiveTestId) { Join-Path $Root 'LocalAppData' } else { [Environment]::GetFolderPath('LocalApplicationData') })
        Registry = $null; Tasks = $tasks
    }
}

function Read-TestJson([string]$root, [string]$name) {
    $p = Join-Path $root $name
    if (Test-Path -LiteralPath $p) { return ([IO.File]::ReadAllText($p) | ConvertFrom-Json) } else { return $null }
}

function Get-TestEnvironment([string]$root) {
    $identity = Read-TestJson $root 'identity.json'
    if (-not $identity) { throw "test mode: identity.json is missing in $root" }
    return [pscustomobject]@{
        Mode = 'test'; UserName = [string]$identity.name; UserSid = [string]$identity.sid
        AppData = (Join-Path $root 'AppData'); LocalAppData = (Join-Path $root 'LocalAppData')
        # Same filter as live mode: only tasks named like ours or running PPTKKeys.dll.
        Registry = (Read-TestJson $root 'registry.json'); Tasks = @(Read-TestJson $root 'tasks.json' | ForEach-Object { $_ } | Where-Object { $_ -and ($_.name -eq $TaskName -or @(@($_.actions) | Where-Object { "$($_.execute) $($_.arguments)" -match 'PPTKKeys\.dll' }).Count -gt 0) })
    }
}

# HKCU reads: the live hive in live mode, registry.json in test mode. Read-only.
function Get-UserValue($envInfo, [string]$key, [string]$name) {
    if ($envInfo.Mode -eq 'test') {
        $r = $envInfo.Registry
        if ((Has $r $key) -and (Has $r.$key $name)) { return $r.$key.$name } else { return $null }
    }
    $p = Get-ItemProperty -LiteralPath ("Registry::HKEY_CURRENT_USER\" + $RegPrefix + $key) -Name $name -ErrorAction SilentlyContinue
    if ($p) { return $p.$name } else { return $null }
}
function Test-UserKey($envInfo, [string]$key) {
    if ($envInfo.Mode -eq 'test') { return (Has $envInfo.Registry $key) }
    return (Test-Path -LiteralPath ("Registry::HKEY_CURRENT_USER\" + $RegPrefix + $key))
}

# A task user id names the current user if it is the SID, the full name, or the bare user name.
function Test-IsCurrentUser([string]$value, $envInfo) {
    if (-not $value) { return $false }
    if ($value -ieq $envInfo.UserSid -or $value -ieq $envInfo.UserName) { return $true }
    $bare = ($envInfo.UserName -split '\\')[-1]
    return ($value -notmatch '\\') -and ($value -ieq $bare)
}

# ---------------------------------------------------------------- manifest
function Read-Manifest {
    $json = $null; $source = ''
    if ($ManifestPath) {
        if (-not (Test-Path -LiteralPath $ManifestPath)) { Problem 'InvalidManifest' "manifest not found: $ManifestPath"; return $null }
        $json = [IO.File]::ReadAllText($ManifestPath); $source = $ManifestPath
    } elseif ($EmbeddedManifestBase64) {
        try { $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($EmbeddedManifestBase64)); $source = 'embedded' } catch { Problem 'InvalidManifest' 'embedded manifest is not valid Base64'; return $null }
    } elseif ($PSScriptRoot -and (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'version.json'))) {
        $json = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'version.json')); $source = 'version.json next to the script'
    } else { Problem 'InvalidManifest' 'no manifest: use the detection script generated by build-package.ps1, or -ManifestPath'; return $null }

    try { $m = $json | ConvertFrom-Json } catch { Problem 'InvalidManifest' "malformed JSON ($source)"; return $null }
    $bad = @()
    if (-not (Has $m 'schemaVersion') -or $m.schemaVersion -ne 1) { $bad += 'schemaVersion must be 1' }
    if (-not (Has $m 'packageName') -or $m.packageName -ne 'PPTK') { $bad += "packageName must be 'PPTK'" }
    if (-not (Has $m 'packageVersion') -or -not (Test-SemVer $m.packageVersion)) { $bad += 'packageVersion missing or not MAJOR.MINOR.PATCH' }
    if (-not (Has $m 'files') -or @($m.files).Count -eq 0) { $bad += 'files missing' }
    else {
        foreach ($f in @($m.files)) {
            if (-not (Has $f 'path') -or -not ($f.path -is [string]) -or $f.path -notmatch '^[A-Za-z0-9._-]+(/[A-Za-z0-9._-]+)*$' -or $f.path -match '(^|/)\.\.?(/|$)') { $bad += "unsafe or missing file path '$(if (Has $f 'path') { $f.path })'"; continue }
            if (-not (Has $f 'sha256') -or -not (Test-Sha256 $f.sha256)) { $bad += "invalid sha256 for $($f.path)" }
        }
        $paths = @($m.files | Where-Object { Has $_ 'path' } | ForEach-Object { $_.path })
        foreach ($req in @('payload/PPTK.ppam') + @($KeysFiles | ForEach-Object { "payload/PPTKKeys/$_" })) { if ($paths -notcontains $req) { $bad += "required entry missing: $req" } }
    }
    if ($bad.Count) { foreach ($b in $bad) { Problem 'InvalidManifest' "$b ($source)" }; return $null }
    return $m
}

# ---------------------------------------------------------------- detection
function Invoke-Detection($envInfo, $m) {
    $version = $m.packageVersion
    $hash = @{}; foreach ($f in @($m.files)) { $hash[$f.path] = $f.sha256 }
    $addinsDir = Join-Path $envInfo.AppData 'Microsoft\AddIns'
    $ppam = Join-Path $addinsDir 'PPTK.ppam'
    $pptkDir = Join-Path $envInfo.LocalAppData 'PPTK'
    $keysDir = Join-Path $pptkDir "versions\$version\PPTKKeys"
    $dll = Join-Path $keysDir 'PPTKKeys.dll'
    $statePath = Join-Path $pptkDir 'state.json'

    $installed = Get-UserValue $envInfo $PptkKey 'InstalledVersion'
    # The installer's two markers. Without both, PPTK is not installed by the managed installer,
    # even if pieces of an earlier manual installation exist (they are listed for diagnosis).
    if (-not $installed -and -not (Test-Path -LiteralPath $statePath)) {
        $legacy = @()
        if (Test-Path -LiteralPath $ppam) { $legacy += 'AddIns\PPTK.ppam' }
        if (Test-UserKey $envInfo $AddinKey) { $legacy += 'add-in registration' }
        if (@($envInfo.Tasks).Count) { $legacy += 'startup task' }
        if (Test-Path -LiteralPath (Join-Path $pptkDir 'versions')) { $legacy += 'PPTK\versions' }
        Problem 'NotInstalled' ("PPTK $version is not installed for $($envInfo.UserName)" + $(if ($legacy.Count) { '; found without installer markers: ' + ($legacy -join ', ') } else { '' }))
        return
    }

    # A: version marker
    if (-not $installed) { Problem 'CorruptState' "HKCU\$PptkKey\InstalledVersion is missing" }
    elseif (-not (Test-SemVer ([string]$installed))) { Problem 'CorruptState' "HKCU\$PptkKey\InstalledVersion '$installed' is not a version" }
    elseif ($installed -ne $version) {
        if ([version]$installed -lt [version]$version) { Problem 'OlderVersion' "installed $installed, expected $version" }
        else { Problem 'WrongVersion' "installed $installed, expected $version" }
    }

    # F: local state (a pending add-in for this version means the update is not applied yet)
    $state = $null
    if (-not (Test-Path -LiteralPath $statePath)) { Problem 'CorruptState' 'state.json is missing' }
    else {
        $state = $null
        try { $state = [IO.File]::ReadAllText($statePath) | ConvertFrom-Json } catch { Problem 'CorruptState' 'state.json is not valid JSON' }
        if ($state) {
            if (-not (Has $state 'installed') -or -not (Test-SemVer $state.installed)) { Problem 'CorruptState' 'state.json has no valid "installed" version' }
            elseif ($state.installed -ne $version) {
                if ([version]$state.installed -lt [version]$version) { Problem 'OlderVersion' "state.json says $($state.installed), expected $version" }
                else { Problem 'WrongVersion' "state.json says $($state.installed), expected $version" }
            }
            if ($installed -and (Has $state 'installed') -and $state.installed -ne $installed) { Problem 'CorruptState' "state.json ($($state.installed)) disagrees with InstalledVersion ($installed)" }
            if ((Has $state 'pending') -and $null -ne $state.pending) {
                if (-not (Has $state.pending 'version') -or -not (Test-SemVer $state.pending.version)) { Problem 'CorruptState' 'state.json pending entry has no valid version' }
                elseif ($state.pending.version -eq $version) { Problem 'PendingUpdate' "PPTK $version is staged but not applied yet (PowerPoint has not been closed)" }
            }
        }
    }

    # B: add-in file
    if (-not (Test-Path -LiteralPath $ppam)) { Problem 'MissingFiles' "add-in missing: AppData\Microsoft\AddIns\PPTK.ppam" }
    else {
        $h = Get-FileSha256 $ppam
        if ($h -ne $hash['payload/PPTK.ppam']) { Problem 'HashMismatch' "PPTK.ppam SHA-256 $h does not match the manifest" }
    }

    # C: PowerPoint add-in registration
    if (-not (Test-UserKey $envInfo $AddinKey)) { Problem 'MissingRegistration' "HKCU\$AddinKey is missing" }
    else {
        $regPath = [string](Get-UserValue $envInfo $AddinKey 'Path')
        $autoLoad = Get-UserValue $envInfo $AddinKey 'AutoLoad'
        $pathOk = ($regPath -ieq 'PPTK.ppam') -or (Same-Path $regPath $ppam)
        if (-not $pathOk) { Problem 'IncorrectRegistration' "add-in Path is '$regPath', expected 'PPTK.ppam' or the AddIns file" }
        if ($null -eq $autoLoad -or [int]$autoLoad -ne 1) { Problem 'IncorrectRegistration' "add-in AutoLoad is '$autoLoad', expected 1" }
    }

    # D: helper files
    foreach ($f in $KeysFiles) {
        $p = Join-Path $keysDir $f
        if (-not (Test-Path -LiteralPath $p)) { Problem 'MissingFiles' "helper file missing: LocalAppData\PPTK\versions\$version\PPTKKeys\$f"; continue }
        $h = Get-FileSha256 $p
        if ($h -ne $hash["payload/PPTKKeys/$f"]) { Problem 'HashMismatch' "$f SHA-256 $h does not match the manifest" }
    }

    # G: private runtime of a universal package
    $runtimeHost = $null
    if ((Has $m 'runtimes') -and $null -ne $m.runtimes) {
        $arch = if ($state -and (Has $state 'architecture')) { [string]$state.architecture } else { '' }
        $native = Get-NativeRid
        if (-not $arch -or -not (Has $m.runtimes $arch)) { Problem 'IncorrectRuntime' "state.json architecture '$arch' is not one of the package's runtimes" }
        elseif ($arch -ne $native) { Problem 'IncorrectRuntime' "installed for $arch, but this device is $native" }
        else {
            $runtimeDir = Join-Path $pptkDir "versions\$version\runtime"
            $runtimeHost = Join-Path $runtimeDir 'dotnet.exe'
            if (-not (Test-Path -LiteralPath $runtimeHost)) { Problem 'MissingFiles' "private runtime missing: LocalAppData\PPTK\versions\$version\runtime" }
            elseif ((Get-RuntimeFingerprint $runtimeDir) -ne [string]$m.runtimes.$arch.fingerprint) { Problem 'HashMismatch' "the private $arch runtime in versions\$version\runtime does not match its fingerprint" }
        }
    }

    # E: startup task
    $tasks = @($envInfo.Tasks)
    $main = @($tasks | Where-Object { $_.name -eq $TaskName -and $_.path -eq $TaskPath })
    $others = @($tasks | Where-Object { -not ($_.name -eq $TaskName -and $_.path -eq $TaskPath) })
    if ($others.Count) { Problem 'IncorrectTask' ("unexpected task(s) running PPTKKeys: " + (($others | ForEach-Object { "$($_.path)$($_.name)" }) -join ', ')) }
    if ($main.Count -eq 0) { Problem 'IncorrectTask' "task '$TaskPath$TaskName' is missing" }
    else {
        $t = $main[0]
        if (-not $t.enabled) { Problem 'IncorrectTask' 'task is disabled' }
        if (-not (Test-IsCurrentUser $t.userId $envInfo)) { Problem 'IncorrectTask' "task principal is '$($t.userId)', expected the current user" }
        if ($t.logonType -ne 'Interactive') { Problem 'IncorrectTask' "task logon type is '$($t.logonType)', expected Interactive" }
        if ($t.runLevel -ne 'Limited') { Problem 'IncorrectTask' "task run level is '$($t.runLevel)', expected Limited" }
        $logon = @(@($t.triggers) | Where-Object { $_.type -eq 'Logon' -and $_.enabled -and (Test-IsCurrentUser $_.userId $envInfo) })
        if ($logon.Count -eq 0) { Problem 'IncorrectTask' 'task has no enabled logon trigger for the current user' }
        $acts = @($t.actions)
        if ($acts.Count -ne 1) { Problem 'IncorrectTask' "task has $($acts.Count) actions, expected 1" }
        else {
            $exe = ([string]$acts[0].execute).Trim('"')
            if ([IO.Path]::GetFileName($exe) -ine 'dotnet.exe') { Problem 'IncorrectTask' "task program is '$exe', expected dotnet.exe" }
            elseif (-not (Test-Path -LiteralPath $exe)) { Problem 'IncorrectTask' "task program '$exe' does not exist" }
            elseif ($runtimeHost -and -not (Same-Path $exe $runtimeHost)) { Problem 'IncorrectTask' "task program is '$exe', expected the private runtime LocalAppData\PPTK\versions\$version\runtime\dotnet.exe" }
            $mm = [regex]::Match([string]$acts[0].arguments, '^\s*exec\s+"([^"]+)"\s*$')
            if (-not $mm.Success) { Problem 'IncorrectTask' "task arguments '$($acts[0].arguments)' are not: exec ""<PPTKKeys.dll>""" }
            elseif (-not (Same-Path $mm.Groups[1].Value $dll)) { Problem 'IncorrectTask' "task runs '$($mm.Groups[1].Value)', expected LocalAppData\PPTK\versions\$version\PPTKKeys\PPTKKeys.dll" }
        }
    }
}

# ---------------------------------------------------------------- main
try {
    $manifest = Read-Manifest
    if ($manifest) {
        $envInfo = if ($Root -and -not $LiveTestId) { Get-TestEnvironment $Root } else { Get-LiveEnvironment }
        Invoke-Detection $envInfo $manifest
    }
} catch {
    Problem 'CorruptState' ("detection error: " + $_.Exception.Message)
}

if ($script:Problems.Count -eq 0) {
    [Console]::Out.WriteLine("PPTK $($manifest.packageVersion) is installed")
    exit 0
}
foreach ($p in $script:Problems) { [Console]::Error.WriteLine($p) }
exit 1
