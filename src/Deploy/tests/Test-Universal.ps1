<#
.SYNOPSIS
  Tests for the universal (x64 + ARM64) package: private .NET runtime selection, staging, verification,
  detection, upgrade from a system-runtime installation, rollback and removal. Isolated test roots only.

.DESCRIPTION
  -RuntimeDir is a folder with the extracted official runtimes: <RuntimeDir>\win-x64\ and <RuntimeDir>\win-arm64\
  (dotnet-runtime-<v>-<rid>.zip, Microsoft-signed). Fake PPTK payloads are combined with those real runtimes, so the
  signature, architecture and framework checks run on real files. The device architecture is set per root
  (<Root>\architecture.txt), so both selections are tested on this machine; no ARM64 binary is executed.
  The live registry, Task Scheduler, profile, PowerPoint and helper are only read (to prove they did not change).
#>
param([Parameter(Mandatory)][string]$RuntimeDir, [string]$RepoRoot = '')

$ErrorActionPreference = 'Stop'
if (-not $RepoRoot) { $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path }
$Install = Join-Path $RepoRoot 'src\Deploy\install.ps1'
$Uninstall = Join-Path $RepoRoot 'src\Deploy\uninstall.ps1'
$Detect = Join-Path $RepoRoot 'src\Deploy\detect.ps1'
$Base = Join-Path ([IO.Path]::GetTempPath()) ('PPTK-universal-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $Base | Out-Null
$Utf8 = New-Object Text.UTF8Encoding($false)
$results = New-Object System.Collections.ArrayList

function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) {
    [void]$results.Add([pscustomobject]@{ Id = $id; Ok = $ok })
    Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail)
}
function Sha([string]$p) { (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLowerInvariant() }
function JsonText($o) { if ($o -is [Array]) { '[' + ((@($o) | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 10 }) -join ',') + ']' } else { ConvertTo-Json -InputObject $o -Depth 10 } }
function Write-Json([string]$p, $o) { [IO.File]::WriteAllText($p, (JsonText $o), $Utf8) }
function Read-Json([string]$p) { if (Test-Path -LiteralPath $p) { [IO.File]::ReadAllText($p) | ConvertFrom-Json | ForEach-Object { $_ } } }
function Machine([string]$p) { $b = [IO.File]::ReadAllBytes($p); '0x{0:X4}' -f [BitConverter]::ToUInt16($b, [BitConverter]::ToInt32($b, 0x3c) + 4) }
function Fingerprint([string]$dir) {
    $dir = (Get-Item -LiteralPath $dir).FullName.TrimEnd('\')
    $lines = [Collections.Generic.List[string]]::new()
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Recurse -File -Force)) { $lines.Add($f.FullName.Substring($dir.Length + 1).Replace('\', '/') + ' ' + (Sha $f.FullName)) }
    $lines.Sort([StringComparer]::Ordinal)
    $h = [Security.Cryptography.SHA256]::Create(); ([BitConverter]::ToString($h.ComputeHash([Text.Encoding]::UTF8.GetBytes(($lines -join "`n") + "`n"))).Replace('-', '')).ToLowerInvariant()
}

# A package: fake add-in and helper files, real runtimes (optionally modified by $tweak before hashing).
function New-Package([string]$name, [string]$version, [string]$ppam, [string]$dll, [switch]$NoRuntime, [scriptblock]$Tweak) {
    $dir = Join-Path $Base "pkg-$name"
    New-Item -ItemType Directory -Force -Path (Join-Path $dir 'payload\PPTKKeys') | Out-Null
    $content = [ordered]@{
        'payload/PPTK.ppam' = "add-in $ppam"
        'payload/PPTKKeys/PPTKKeys.dll' = "helper $dll"
        'payload/PPTKKeys/PPTKKeys.deps.json' = '{"runtimeTarget":{"name":".NETCoreApp,Version=v10.0"}}'
        'payload/PPTKKeys/PPTKKeys.runtimeconfig.json' = '{"runtimeOptions":{"tfm":"net10.0","framework":{"name":"Microsoft.NETCore.App","version":"10.0.0"}}}'
    }
    foreach ($k in $content.Keys) { [IO.File]::WriteAllText((Join-Path $dir ($k.Replace('/', '\'))), $content[$k], $Utf8) }
    $runtimes = [ordered]@{}
    if (-not $NoRuntime) {
        foreach ($rid in 'win-x64', 'win-arm64') {
            $dst = Join-Path $dir "payload\runtime\$rid"; New-Item -ItemType Directory -Force -Path $dst | Out-Null
            Copy-Item -Path (Join-Path $RuntimeDir "$rid\*") -Destination $dst -Recurse
        }
        if ($Tweak) { & $Tweak $dir }
        foreach ($rid in 'win-x64', 'win-arm64') { $runtimes[$rid] = [ordered]@{ version = '10.0.15'; fingerprint = (Fingerprint (Join-Path $dir "payload\runtime\$rid")) } }
    }
    $files = @(Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.Name -ne 'version.json' } | ForEach-Object { [ordered]@{ path = $_.FullName.Substring($dir.Length + 1).Replace('\', '/'); sha256 = (Sha $_.FullName) } })
    $m = [ordered]@{ schemaVersion = 1; packageName = 'PPTK'; packageVersion = $version; files = $files }
    if ($runtimes.Count) { $m['runtimes'] = $runtimes }
    Write-Json (Join-Path $dir 'version.json') $m
    return $dir
}

function New-Root([string]$name, [string]$arch = 'win-x64') {
    $r = Join-Path $Base $name
    New-Item -ItemType Directory -Force -Path $r | Out-Null
    [IO.File]::WriteAllText((Join-Path $r 'PPTK-TEST-ROOT'), 'isolated PPTK universal test root', $Utf8)
    Write-Json (Join-Path $r 'identity.json') ([ordered]@{ name = 'TESTDOM\tester'; sid = 'S-1-5-21-1000-2000-3000-1001' })
    Write-Json (Join-Path $r 'registry.json') ([ordered]@{ 'Software\VB and VBA Program Settings\Instrumenta\General' = [ordered]@{ OperatingMode = 'pro' } })
    Write-Json (Join-Path $r 'processes.json') @()
    [IO.File]::WriteAllText((Join-Path $r 'architecture.txt'), $arch, $Utf8)
    # A system .NET for the old-style (1.0.0) package; never executed.
    $dn = Join-Path $r 'LocalAppData\Microsoft\dotnet'
    New-Item -ItemType Directory -Force -Path (Join-Path $dn 'shared\Microsoft.NETCore.App\10.0.14') | Out-Null
    [IO.File]::WriteAllText((Join-Path $dn 'dotnet.exe'), 'placeholder; never executed', $Utf8)
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
function Install([string]$root, [string]$pkg) { Invoke-Ps $Install @('-Root', $root, '-PackagePath', $pkg) }
function Rollback([string]$root) { Invoke-Ps $Install @('-Root', $root, '-Rollback') }
function Uninstall([string]$root) { Invoke-Ps $Uninstall @('-Root', $root) }
function Detect([string]$root, [string]$pkg) { Invoke-Ps $Detect @('-Root', $root, '-ManifestPath', (Join-Path $pkg 'version.json')) }
function Detected([string]$root, [string]$pkg) { $d = Detect $root $pkg; return ($d.Code -eq 0 -and $d.Out -match 'is installed') }
function Short($r) { "exit $($r.Code): " + (($r.Text -split "`r?`n")[0]) }
function State([string]$root) { Read-Json (Join-Path $root 'LocalAppData\PPTK\state.json') }
function Task([string]$root) { @(Read-Json (Join-Path $root 'tasks.json') | Where-Object { $_.name -eq 'PPTKKeys Startup' })[0] }
function Host-Of([string]$root) { [string](Task $root).actions[0].execute }
function Snap([string]$root) { @(Get-ChildItem -LiteralPath $root -Recurse -File -Force | Where-Object { $_.FullName -notmatch '\\logs\\|state\.lock$|helper-actions\.log$' } | ForEach-Object { $_.FullName.Substring($root.Length) + '|' + (Sha $_.FullName) }) -join "`n" }
function Live-Snap {
    @(
        'task=' + $(try { [string](Export-ScheduledTask -TaskName 'PPTKKeys Startup' -ErrorAction Stop) } catch { '' })
        'pptkKey=' + (Test-Path 'HKCU:\Software\PPTK')
        'local=' + $(if (Test-Path (Join-Path $env:LOCALAPPDATA 'PPTK')) { @(Get-ChildItem (Join-Path $env:LOCALAPPDATA 'PPTK') -Recurse -Force -File | ForEach-Object { $_.FullName + '|' + $_.Length + '|' + $_.LastWriteTimeUtc.Ticks }) -join ';' } else { '' })
        'helper=' + (@(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'dotnet.exe' -and $_.CommandLine -like '*PPTKKeys.dll*' } | ForEach-Object { $_.ProcessId }) -join ',')
    ) -join "`n"
}

$liveBefore = Live-Snap
try {
    foreach ($rid in 'win-x64', 'win-arm64') { if (-not (Test-Path -LiteralPath (Join-Path $RuntimeDir "$rid\dotnet.exe"))) { throw "missing $RuntimeDir\$rid\dotnet.exe" } }
    $P100 = New-Package 'v100' '1.0.0' 'A' 'A' -NoRuntime          # the current release: system .NET
    $P110 = New-Package 'v110' '1.1.0' 'B' 'A'                     # universal, helper DLL unchanged
    $P111 = New-Package 'v111' '1.1.1' 'C' 'C'                     # universal, helper changed

    # ---------------------------------------------------------------- selection by device architecture
    $r = New-Root 'X64' 'win-x64'; $i = Install $r $P110; $st = State $r
    $hx = Join-Path $r 'LocalAppData\PPTK\versions\1.1.0\runtime\dotnet.exe'
    Record 'U1' 'x64 device: x64 runtime staged, task runs it, detected' ($i.Code -eq 0 -and (Detected $r $P110) -and (Machine $hx) -eq '0x8664' -and (Host-Of $r) -eq $hx -and $st.architecture -eq 'win-x64') "$(Short $i); host $(Machine $hx)"

    $r = New-Root 'ARM64' 'win-arm64'; $i = Install $r $P110; $st = State $r
    $ha = Join-Path $r 'LocalAppData\PPTK\versions\1.1.0\runtime\dotnet.exe'
    $onlyOne = -not (Test-Path (Join-Path $r 'LocalAppData\PPTK\versions\1.1.0\runtime\win-x64'))
    Record 'U2' 'ARM64 device: ARM64 runtime staged (and only that one), task runs it, detected' ($i.Code -eq 0 -and (Detected $r $P110) -and (Machine $ha) -eq '0xAA64' -and (Host-Of $r) -eq $ha -and $st.architecture -eq 'win-arm64' -and $onlyOne) "$(Short $i); host $(Machine $ha)"

    [IO.File]::WriteAllText((Join-Path $r 'architecture.txt'), 'win-x64', $Utf8)
    $d = Detect $r $P110
    Record 'U3' 'installed for ARM64 but the device reports x64: not detected (IncorrectRuntime)' ($d.Code -eq 1 -and $d.Err -match 'IncorrectRuntime') (($d.Err -split "`n")[0])

    $r = New-Root 'Unsupported' 'unsupported (processor architecture 0 / x86)'; $before = Snap $r; $i = Install $r $P110
    Record 'U4' 'unsupported architecture: refused, nothing changed' ($i.Code -eq 1 -and $i.Err -match 'UnsupportedArchitecture' -and (Snap $r) -eq $before) (Short $i)

    # ---------------------------------------------------------------- integrity of the runtime
    $r = New-Root 'Tamper'; [void](Install $r $P110)
    [IO.File]::AppendAllText((Join-Path $r 'LocalAppData\PPTK\versions\1.1.0\runtime\shared\Microsoft.NETCore.App\10.0.15\System.Private.CoreLib.dll'), 'x')
    $d = Detect $r $P110
    Record 'U5' 'changed runtime file after installation: not detected (HashMismatch)' ($d.Code -eq 1 -and $d.Err -match 'HashMismatch.*runtime') (($d.Err -split "`n")[0])
    Remove-Item -LiteralPath (Join-Path $r 'LocalAppData\PPTK\versions\1.1.0\runtime\host') -Recurse -Force
    $i = Install $r $P110
    Record 'U6' 'damaged runtime is repaired by installing again' ($i.Code -eq 0 -and (Detected $r $P110)) (Short $i)

    # Same-version repair replaces versions\<v>, the folder the running helper uses (its files are locked):
    # the helper is asked to close before the folder is replaced and started again afterwards.
    $r = New-Root 'SameVersion'; [void](Install $r $P110)
    Remove-Item -LiteralPath (Join-Path $r 'helper-actions.log') -ErrorAction SilentlyContinue
    [IO.File]::AppendAllText((Join-Path $r 'LocalAppData\PPTK\versions\1.1.0\PPTKKeys\PPTKKeys.dll'), 'x')
    $i = Install $r $P110
    $acts = @(Get-Content (Join-Path $r 'helper-actions.log') -ErrorAction SilentlyContinue) -join ','
    Record 'U6b' 'same-version repair of the helper folder: helper closed first, started again, detected' ($i.Code -eq 0 -and (Detected $r $P110) -and $acts -eq 'stop,start') "$(Short $i); actions $acts"

    $Punsigned = New-Package 'unsigned' '1.1.0' 'B' 'A' -Tweak { param($dir) [IO.File]::WriteAllText((Join-Path $dir 'payload\runtime\win-x64\shared\Microsoft.NETCore.App\10.0.15\Extra.dll'), 'not a signed binary') }
    $r = New-Root 'Unsigned'; $before = Snap $r; $i = Install $r $Punsigned
    Record 'U7' 'runtime with a file not signed by Microsoft: refused (InvalidRuntime), nothing installed' ($i.Code -eq 1 -and $i.Err -match 'InvalidRuntime' -and -not (Test-Path (Join-Path $r 'LocalAppData\PPTK\versions\1.1.0')) -and -not (Task $r)) (Short $i)

    $Pswapped = New-Package 'swapped' '1.1.0' 'B' 'A' -Tweak { param($dir) Remove-Item (Join-Path $dir 'payload\runtime\win-arm64') -Recurse -Force; Copy-Item (Join-Path $RuntimeDir 'win-x64') (Join-Path $dir 'payload\runtime\win-arm64') -Recurse }
    $r = New-Root 'Swapped' 'win-arm64'; $i = Install $r $Pswapped
    Record 'U8' 'x64 binaries offered as the ARM64 runtime: refused (InvalidRuntime), nothing installed' ($i.Code -eq 1 -and $i.Err -match 'InvalidRuntime.*machine' -and -not (Task $r)) (Short $i)

    $Pexe = New-Package 'exe' '1.1.0' 'B' 'A' -Tweak { param($dir) Copy-Item (Join-Path $RuntimeDir 'win-x64\dotnet.exe') (Join-Path $dir 'payload\PPTKKeys\tool.exe') }
    $r = New-Root 'Exe'; $i = Install $r $Pexe
    Record 'U9' 'an executable outside the runtimes (even Microsoft-signed): refused' ($i.Code -eq 1 -and $i.Err -match 'banned file type.*tool\.exe' -and -not (Task $r)) (Short $i)

    # ---------------------------------------------------------------- upgrade, pending, rollback, removal
    $r = New-Root 'Upgrade'; [void](Install $r $P100)
    $sysHost = Host-Of $r
    $i = Install $r $P110; $st = State $r
    $h110 = Join-Path $r 'LocalAppData\PPTK\versions\1.1.0\runtime\dotnet.exe'
    $restart = @(Get-Content (Join-Path $r 'helper-actions.log')) -join ','
    Record 'U10' '1.0.0 (system .NET) -> 1.1.0 universal: task switches to the private runtime, helper restarted' ($i.Code -eq 0 -and (Detected $r $P110) -and (Host-Of $r) -eq $h110 -and $st.previous -eq '1.0.0' -and $restart -match 'stop,start$') "$(Short $i); actions $restart"

    $i = Rollback $r
    Record 'U11' 'rollback 1.1.0 -> 1.0.0: back on the system .NET, detected' ($i.Code -eq 0 -and (Detected $r $P100) -and (Host-Of $r) -eq $sysHost) (Short $i)

    $i = Install $r $P110
    Write-Json (Join-Path $r 'processes.json') @('POWERPNT')
    $i2 = Install $r $P111; $st = State $r
    Record 'U12' 'PowerPoint open: 1.1.1 staged with its runtime; pending keysChanged, architecture recorded; 1.1.0 still detected' ($i.Code -eq 0 -and $i2.Code -eq 1618 -and $st.pending.version -eq '1.1.1' -and $st.pending.keysChanged -eq $true -and $st.architecture -eq 'win-x64' -and (Test-Path (Join-Path $r 'LocalAppData\PPTK\versions\1.1.1\runtime\dotnet.exe')) -and (Detected $r $P110)) "$(Short $i) / $(Short $i2)"

    Write-Json (Join-Path $r 'processes.json') @()
    $i = Install $r $P111
    $h111 = Join-Path $r 'LocalAppData\PPTK\versions\1.1.1\runtime\dotnet.exe'
    Record 'U13' '1.1.0 -> 1.1.1 universal: task moves to the 1.1.1 runtime; 1.1.0 kept for rollback' ($i.Code -eq 0 -and (Detected $r $P111) -and (Host-Of $r) -eq $h111 -and (Test-Path $h110)) (Short $i)

    $i = Rollback $r
    Record 'U14' 'rollback 1.1.1 -> 1.1.0: task back on the 1.1.0 private runtime' ($i.Code -eq 0 -and (Detected $r $P110) -and (Host-Of $r) -eq $h110) (Short $i)

    $i = Uninstall $r
    $d = Detect $r $P110
    Record 'U15' 'uninstall removes the versions with their runtimes, the task and the state' ($i.Code -eq 0 -and -not (Test-Path (Join-Path $r 'LocalAppData\PPTK\versions')) -and -not (Task $r) -and $d.Err -match 'NotInstalled') (Short $i)

    $r = New-Root 'Fail'; [void](Install $r $P100); $hostBefore = Host-Of $r; $taskBefore = JsonText (Task $r)
    $i = Invoke-Ps $Install @('-Root', $r, '-PackagePath', $P110, '-TestFailAt', 'addin')
    $st = State $r; $reg = Read-Json (Join-Path $r 'registry.json')
    # state.json keeps lastError (by design); everything else is as before.
    Record 'U16' 'failed upgrade to universal: task, registry and state restored, the staged runtime removed' ($i.Code -eq 1 -and (Detected $r $P100) -and (JsonText (Task $r)) -eq $taskBefore -and (Host-Of $r) -eq $hostBefore -and $st.installed -eq '1.0.0' -and -not $st.pending -and $st.lastError -match 'simulated' -and $reg.'Software\PPTK'.InstalledVersion -eq '1.0.0' -and -not (Test-Path (Join-Path $r 'LocalAppData\PPTK\versions\1.1.0'))) (Short $i)
} finally {
    Remove-Item -LiteralPath $Base -Recurse -Force -ErrorAction SilentlyContinue
}
Record 'U17' 'live task, keys, PPTK folder and helper unchanged' ((Live-Snap) -eq $liveBefore) ''

$failed = @($results | Where-Object { -not $_.Ok }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
