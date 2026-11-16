<#
.SYNOPSIS
  Tests for src\Deploy\detect.ps1 (Phase 1A).

.DESCRIPTION
  Every case builds an isolated installation under a temporary -Root (files, registry.json,
  tasks.json, identity.json), changes one thing, and runs detect.ps1 in a separate PowerShell
  process. A case passes when the exit code, STDOUT and the STDERR category are as expected.
  The live registry, Task Scheduler and profile are only read, by the live cases (R, T, Z), and
  are compared before and after. The "dotnet.exe" in a test root is an empty text file used only
  for the path check; nothing is executed.

  With -Package <folder> (a package built by build-package.ps1), case Z also runs the generated
  PPTK-detect.ps1 next to it: against a test root holding that package's real files (detected) and
  live (read-only).

  Exit code 0 when every case passes, 1 otherwise.
#>
param([string]$RepoRoot = '', [string]$Package = '')

$ErrorActionPreference = 'Stop'
if (-not $RepoRoot) { $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path }
$Detect = Join-Path $RepoRoot 'src\Deploy\detect.ps1'
$Base = Join-Path ([IO.Path]::GetTempPath()) ('PPTK-detect-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $Base | Out-Null
$results = New-Object System.Collections.ArrayList
$AddinKey = 'Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK'
$Utf8 = New-Object Text.UTF8Encoding($false)

function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) {
    [void]$results.Add([pscustomobject]@{ Id = $id; Ok = $ok })
    Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail)
}
# Arrays are written element by element: in Windows PowerShell 5.1 an array that came from
# ConvertFrom-Json serializes as {"value": [...], "Count": n}.
function Write-Json([string]$path, $obj) {
    if ($obj -is [Array]) { $json = '[' + ((@($obj) | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 8 }) -join ',') + ']' }
    else { $json = ConvertTo-Json -InputObject $obj -Depth 8 }
    [IO.File]::WriteAllText($path, $json, $Utf8)
}
function Sha([string]$path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() }

# A fake package: small files with known content and a valid version.json.
function New-FakePackage([string]$dir, [string]$version = '1.0.0') {
    New-Item -ItemType Directory -Force -Path (Join-Path $dir 'payload\PPTKKeys') | Out-Null
    $files = [ordered]@{
        'payload/PPTK.ppam' = "ppam $version"
        'payload/PPTKKeys/PPTKKeys.dll' = "dll $version"
        'payload/PPTKKeys/PPTKKeys.deps.json' = '{"deps":1}'
        'payload/PPTKKeys/PPTKKeys.runtimeconfig.json' = '{"runtime":1}'
    }
    $list = @()
    foreach ($k in $files.Keys) {
        $p = Join-Path $dir ($k.Replace('/', '\')); [IO.File]::WriteAllText($p, $files[$k], $Utf8)
        $list += [ordered]@{ path = $k; size = (Get-Item -LiteralPath $p).Length; sha256 = (Sha $p) }
    }
    Write-Json (Join-Path $dir 'version.json') ([ordered]@{ schemaVersion = 1; packageName = 'PPTK'; packageVersion = $version; files = $list })
    return (Join-Path $dir 'version.json')
}

# A complete, healthy installation of the package in $pkgDir under $root.
function New-Install([string]$root, [string]$pkgDir, [string]$version = '1.0.0') {
    $addins = Join-Path $root 'AppData\Microsoft\AddIns'
    $keys = Join-Path $root "LocalAppData\PPTK\versions\$version\PPTKKeys"
    $dotnet = Join-Path $root 'dotnet'
    New-Item -ItemType Directory -Force -Path $addins, $keys, $dotnet | Out-Null
    Copy-Item -LiteralPath (Join-Path $pkgDir 'payload\PPTK.ppam') -Destination $addins
    foreach ($f in 'PPTKKeys.dll', 'PPTKKeys.deps.json', 'PPTKKeys.runtimeconfig.json') { Copy-Item -LiteralPath (Join-Path $pkgDir "payload\PPTKKeys\$f") -Destination $keys }
    [IO.File]::WriteAllText((Join-Path $dotnet 'dotnet.exe'), 'placeholder for the path check; never executed', $Utf8)
    Write-Json (Join-Path $root 'identity.json') ([ordered]@{ name = 'TESTDOM\tester'; sid = 'S-1-5-21-1000-2000-3000-1001' })
    Write-Json (Join-Path $root 'registry.json') ([ordered]@{ 'Software\PPTK' = [ordered]@{ InstalledVersion = $version }; $AddinKey = [ordered]@{ Path = 'PPTK.ppam'; AutoLoad = 1 } })
    Write-Json (Join-Path $root 'LocalAppData\PPTK\state.json') ([ordered]@{ installed = $version; pending = $null; lastError = $null })
    Write-Json (Join-Path $root 'tasks.json') @([ordered]@{
        name = 'PPTKKeys Startup'; path = '\'; enabled = $true; userId = 'tester'; logonType = 'Interactive'; runLevel = 'Limited'
        triggers = @([ordered]@{ type = 'Logon'; userId = 'TESTDOM\tester'; enabled = $true })
        actions = @([ordered]@{ execute = (Join-Path $dotnet 'dotnet.exe'); arguments = 'exec "' + (Join-Path $keys 'PPTKKeys.dll') + '"' })
    })
}

function Invoke-Detect([string]$script, [string[]]$arguments) {
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $script + '" ' + (($arguments | ForEach-Object { '"' + $_ + '"' }) -join ' ')
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8; $psi.StandardErrorEncoding = [Text.Encoding]::UTF8
    $pr = [Diagnostics.Process]::Start($psi)
    $out = $pr.StandardOutput.ReadToEnd(); $err = $pr.StandardError.ReadToEnd(); $pr.WaitForExit()
    return [pscustomobject]@{ Code = $pr.ExitCode; Out = $out.Trim(); Err = $err.Trim() }
}

$script:case = 0
function New-Case { $script:case++; $r = Join-Path $Base ("case{0:D2}" -f $script:case); New-Install $r $PkgDir; return $r }
function Run-Root([string]$root, [string]$manifest = $Manifest) { Invoke-Detect $Detect @('-Root', $root, '-ManifestPath', $manifest) }
function Expect-Detected([string]$id, [string]$name, $r) {
    Record $id $name ($r.Code -eq 0 -and $r.Out -eq 'PPTK 1.0.0 is installed' -and $r.Err -eq '') ("exit $($r.Code); stdout '$($r.Out)'" + $(if ($r.Err) { "; stderr '$($r.Err)'" }))
}
function Expect-Not([string]$id, [string]$name, $r, [string]$category, [string]$text = '') {
    $first = ($r.Err -split "`r?`n")[0]
    $ok = ($r.Code -ne 0) -and ($r.Out -eq '') -and ($r.Err -match "PPTK-DETECT ${category}:") -and (-not $text -or $r.Err -match $text)
    Record $id $name $ok ("exit $($r.Code); " + $(if ($r.Out) { "STDOUT '$($r.Out)'; " }) + $first)
}
function Edit-Json([string]$path, [scriptblock]$change) { $o = [IO.File]::ReadAllText($path) | ConvertFrom-Json; & $change $o; Write-Json $path $o }

try {
    $PkgDir = Join-Path $Base 'pkg'; $Manifest = New-FakePackage $PkgDir

    $r = New-Case; Expect-Detected 'A' 'fully valid installation' (Run-Root $r)

    $r = Join-Path $Base 'emptyB'; New-Item -ItemType Directory -Path $r | Out-Null
    Write-Json (Join-Path $r 'identity.json') ([ordered]@{ name = 'TESTDOM\tester'; sid = 'S-1-5-21-1' })
    Expect-Not 'B' 'no installation' (Run-Root $r) 'NotInstalled'

    $r = New-Case; Edit-Json (Join-Path $r 'registry.json') { param($o) $o.'Software\PPTK'.InstalledVersion = '0.9.0' }; Edit-Json (Join-Path $r 'LocalAppData\PPTK\state.json') { param($o) $o.installed = '0.9.0' }
    Expect-Not 'C' 'older version installed' (Run-Root $r) 'OlderVersion' 'installed 0\.9\.0, expected 1\.0\.0'

    $r = New-Case; Edit-Json (Join-Path $r 'registry.json') { param($o) $o.'Software\PPTK'.InstalledVersion = '1.1.0' }; Edit-Json (Join-Path $r 'LocalAppData\PPTK\state.json') { param($o) $o.installed = '1.1.0' }
    Expect-Not 'D1' 'newer/different version installed' (Run-Root $r) 'WrongVersion' 'installed 1\.1\.0'
    $r = New-Case; Edit-Json (Join-Path $r 'registry.json') { param($o) $o.'Software\PPTK'.InstalledVersion = 'one point oh' }
    Expect-Not 'D2' 'invalid installed version value' (Run-Root $r) 'CorruptState' 'is not a version'

    $r = New-Case; Remove-Item -LiteralPath (Join-Path $r 'AppData\Microsoft\AddIns\PPTK.ppam')
    Expect-Not 'E' 'PPAM missing' (Run-Root $r) 'MissingFiles' 'AddIns\\PPTK\.ppam'
    $r = New-Case; [IO.File]::AppendAllText((Join-Path $r 'AppData\Microsoft\AddIns\PPTK.ppam'), 'x')
    Expect-Not 'F' 'PPAM hash mismatch' (Run-Root $r) 'HashMismatch' 'PPTK\.ppam SHA-256'

    $kd = 'LocalAppData\PPTK\versions\1.0.0\PPTKKeys'
    $r = New-Case; Remove-Item -LiteralPath (Join-Path $r "$kd\PPTKKeys.dll")
    Expect-Not 'G' 'PPTKKeys DLL missing' (Run-Root $r) 'MissingFiles' 'PPTKKeys\.dll'
    $r = New-Case; [IO.File]::AppendAllText((Join-Path $r "$kd\PPTKKeys.dll"), 'x')
    Expect-Not 'H' 'PPTKKeys hash mismatch' (Run-Root $r) 'HashMismatch' 'PPTKKeys\.dll SHA-256'
    $r = New-Case; Remove-Item -LiteralPath (Join-Path $r "$kd\PPTKKeys.deps.json"), (Join-Path $r "$kd\PPTKKeys.runtimeconfig.json")
    $res = Run-Root $r
    Expect-Not 'I' 'dependency JSON files missing' $res 'MissingFiles' 'deps\.json'
    Record 'I2' 'both missing dependency files are reported' ($res.Err -match 'runtimeconfig\.json' -and $res.Err -match 'deps\.json') ''

    $r = New-Case; Edit-Json (Join-Path $r 'registry.json') { param($o) $o.PSObject.Properties.Remove('Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK') }
    Expect-Not 'J' 'PowerPoint registration missing' (Run-Root $r) 'MissingRegistration'
    $r = New-Case; Edit-Json (Join-Path $r 'registry.json') { param($o) $o.'Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK'.Path = 'C:\Somewhere\Other.ppam' }
    Expect-Not 'K1' 'registration Path points elsewhere' (Run-Root $r) 'IncorrectRegistration' 'Path'
    $r = New-Case; Edit-Json (Join-Path $r 'registry.json') { param($o) $o.'Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK'.AutoLoad = 0 }
    Expect-Not 'K2' 'registration AutoLoad 0' (Run-Root $r) 'IncorrectRegistration' 'AutoLoad'
    $r = New-Case; Edit-Json (Join-Path $r 'registry.json') { param($o) $o.'Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK'.Path = (Join-Path $r 'AppData\Microsoft\AddIns\PPTK.ppam') }
    Expect-Detected 'K3' 'registration with the full AddIns path is accepted' (Run-Root $r)

    $r = New-Case; Write-Json (Join-Path $r 'tasks.json') @()
    Expect-Not 'L' 'scheduled task missing' (Run-Root $r) 'IncorrectTask' 'is missing'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].userId = 'OTHERDOM\someoneelse' }
    Expect-Not 'M1' 'task principal is another user' (Run-Root $r) 'IncorrectTask' 'principal'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].triggers[0].userId = 'TESTDOM\other' }
    Expect-Not 'M2' 'logon trigger for another user' (Run-Root $r) 'IncorrectTask' 'logon trigger'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].triggers[0].type = 'Boot' }
    Expect-Not 'M3' 'wrong trigger type' (Run-Root $r) 'IncorrectTask' 'logon trigger'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].logonType = 'Password' }
    Expect-Not 'M4' 'not interactive' (Run-Root $r) 'IncorrectTask' 'logon type'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].runLevel = 'Highest' }
    Expect-Not 'M5' 'highest privileges' (Run-Root $r) 'IncorrectTask' 'run level'
    $r = New-Case; $t = [IO.File]::ReadAllText((Join-Path $r 'tasks.json')) | ConvertFrom-Json; $dup = $t[0].PSObject.Copy(); $dup.name = 'PPTKKeys Copy'
    Write-Json (Join-Path $r 'tasks.json') @($t[0], $dup)
    Expect-Not 'M6' 'duplicate task running PPTKKeys' (Run-Root $r) 'IncorrectTask' 'unexpected task'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].enabled = $false }
    Expect-Not 'M7' 'task disabled' (Run-Root $r) 'IncorrectTask' 'disabled'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].userId = 'S-1-5-21-1000-2000-3000-1001' }
    Expect-Detected 'M8' 'principal given as the user SID is accepted' (Run-Root $r)

    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].actions[0].execute = (Join-Path $r 'dotnet\missing\dotnet.exe') }
    Expect-Not 'N1' 'dotnet.exe path does not exist' (Run-Root $r) 'IncorrectTask' 'does not exist'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].actions[0].execute = 'C:\Windows\System32\cmd.exe' }
    Expect-Not 'N2' 'task host is not dotnet.exe' (Run-Root $r) 'IncorrectTask' 'expected dotnet\.exe'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].actions[0].arguments = 'exec "' + (Join-Path $r 'LocalAppData\PPTK\PPTKKeys\PPTKKeys.dll') + '"' }
    Expect-Not 'N3' 'task runs the DLL from another folder (old layout)' (Run-Root $r) 'IncorrectTask' 'expected LocalAppData'
    $r = New-Case; Edit-Json (Join-Path $r 'tasks.json') { param($o) $o[0].actions[0].arguments = '"' + (Join-Path $r "$kd\PPTKKeys.dll") + '"' }
    Expect-Not 'N4' 'task arguments without exec' (Run-Root $r) 'IncorrectTask' 'are not'

    $r = New-Case; Edit-Json (Join-Path $r 'LocalAppData\PPTK\state.json') { param($o) $o.pending = [ordered]@{ version = '1.0.0'; file = 'pending\PPTK-1.0.0.ppam'; sha256 = ('a' * 64); notified = $true } }
    Expect-Not 'O1' 'pending update for this version (files already match)' (Run-Root $r) 'PendingUpdate'
    $r = New-Case; Edit-Json (Join-Path $r 'LocalAppData\PPTK\state.json') { param($o) $o.pending = [ordered]@{ version = '1.0.0'; file = 'pending\PPTK-1.0.0.ppam'; sha256 = ('a' * 64); notified = $false } }
    [IO.File]::WriteAllText((Join-Path $r 'AppData\Microsoft\AddIns\PPTK.ppam'), 'previous add-in', $Utf8)
    $res = Run-Root $r
    Expect-Not 'O2' 'pending update with the old add-in still in place' $res 'PendingUpdate'
    Record 'O3' 'pending case also reports the add-in hash' ($res.Err -match 'HashMismatch') ''

    $r = New-Case; Remove-Item -LiteralPath (Join-Path $r 'LocalAppData\PPTK\state.json')
    Expect-Not 'P1' 'state.json missing' (Run-Root $r) 'CorruptState' 'state\.json is missing'
    $r = New-Case; [IO.File]::WriteAllText((Join-Path $r 'LocalAppData\PPTK\state.json'), '{ "installed": "1.0.0", ', $Utf8)
    Expect-Not 'P2' 'state.json malformed' (Run-Root $r) 'CorruptState' 'not valid JSON'
    $r = New-Case; Edit-Json (Join-Path $r 'LocalAppData\PPTK\state.json') { param($o) $o.installed = '1.0.1' }
    Expect-Not 'P3' 'state.json disagrees with InstalledVersion' (Run-Root $r) 'CorruptState' 'disagrees'
    $r = New-Case; Edit-Json (Join-Path $r 'LocalAppData\PPTK\state.json') { param($o) $o.PSObject.Properties.Remove('installed') }
    Expect-Not 'P4' 'state.json without "installed"' (Run-Root $r) 'CorruptState' 'no valid "installed"'
    $r = New-Case; Edit-Json (Join-Path $r 'registry.json') { param($o) $o.'Software\PPTK'.PSObject.Properties.Remove('InstalledVersion') }
    Expect-Not 'P5' 'InstalledVersion missing but state.json present' (Run-Root $r) 'CorruptState' 'InstalledVersion is missing'

    # Q: manifest validation
    $r = New-Case
    Expect-Not 'Q1' 'manifest missing' (Invoke-Detect $Detect @('-Root', $r, '-ManifestPath', (Join-Path $Base 'nope.json'))) 'InvalidManifest' 'not found'
    $bad = Join-Path $Base 'bad1.json'; [IO.File]::WriteAllText($bad, '{ "packageName": ', $Utf8)
    Expect-Not 'Q2' 'manifest malformed JSON' (Run-Root $r $bad) 'InvalidManifest' 'malformed'
    $mk = { param($name, [scriptblock]$change) $o = [IO.File]::ReadAllText($Manifest) | ConvertFrom-Json; & $change $o; $p = Join-Path $Base $name; Write-Json $p $o; $p }
    Expect-Not 'Q3' 'manifest version invalid' (Run-Root $r (& $mk 'q3.json' { param($o) $o.packageVersion = '1.0' })) 'InvalidManifest' 'packageVersion'
    Expect-Not 'Q4' 'manifest SHA-256 invalid' (Run-Root $r (& $mk 'q4.json' { param($o) $o.files[1].sha256 = 'XYZ' })) 'InvalidManifest' 'invalid sha256'
    Expect-Not 'Q5' 'manifest unsafe relative path' (Run-Root $r (& $mk 'q5.json' { param($o) $o.files[0].path = '../../Windows/evil.ppam' })) 'InvalidManifest' 'unsafe'
    Expect-Not 'Q6' 'manifest absolute path' (Run-Root $r (& $mk 'q6.json' { param($o) $o.files[0].path = 'C:/x/PPTK.ppam' })) 'InvalidManifest' 'unsafe'
    Expect-Not 'Q7' 'manifest missing required entry' (Run-Root $r (& $mk 'q7.json' { param($o) $o.files = @($o.files | Where-Object { $_.path -ne 'payload/PPTKKeys/PPTKKeys.deps.json' }) })) 'InvalidManifest' 'required entry missing'
    Expect-Not 'Q8' 'manifest without packageName' (Run-Root $r (& $mk 'q8.json' { param($o) $o.PSObject.Properties.Remove('packageName') })) 'InvalidManifest' 'packageName'

    # R: repeated detection has no side effects (test root and live state)
    $r = New-Case
    # Every path (catches anything created or deleted) and, for files, the content hash and the
    # write time read from the file itself. Folder times from a directory listing are left out:
    # NTFS refreshes that cached copy in the parent index lazily, even when nothing is written.
    $snap = { param($dir) @(Get-ChildItem -LiteralPath $dir -Recurse -Force | ForEach-Object {
        if ($_.PSIsContainer) { 'D ' + $_.FullName } else { 'F ' + $_.FullName + '|' + (Get-Item -LiteralPath $_.FullName).LastWriteTimeUtc.Ticks + '|' + (Sha $_.FullName) }
    }) -join "`n" }
    $before = & $snap $r; $x1 = Run-Root $r; $x2 = Run-Root $r; $after = & $snap $r
    $diff = @(Compare-Object @($before -split "`n") @($after -split "`n") | ForEach-Object { $_.SideIndicator + ' ' + $_.InputObject })
    Record 'R1' 'repeated detection: same result, test root unchanged' ($before -eq $after -and $x1.Code -eq 0 -and $x2.Code -eq 0 -and $x1.Out -eq $x2.Out) ("runs exit $($x1.Code)/$($x2.Code); root changes: $($diff.Count) " + ($diff | Select-Object -First 2) + " " + (($x1.Err -split "`r?`n")[0]))
    $live = { [pscustomobject]@{
        Task = $(try { Export-ScheduledTask -TaskName 'PPTKKeys Startup' -ErrorAction Stop } catch { '' })
        Pptk = (Test-Path 'HKCU:\Software\PPTK')
        Reg = ((Get-ItemProperty -Path ('HKCU:\' + $AddinKey) -ErrorAction SilentlyContinue | Select-Object Path, AutoLoad | Out-String))
        Local = $(if (Test-Path (Join-Path $env:LOCALAPPDATA 'PPTK')) { @(Get-ChildItem (Join-Path $env:LOCALAPPDATA 'PPTK') -Recurse -Force | ForEach-Object { $_.FullName + '|' + $_.LastWriteTimeUtc.Ticks }) -join ';' } else { '' }) } }
    $lb = & $live; $l1 = Invoke-Detect $Detect @('-ManifestPath', $Manifest); $l2 = Invoke-Detect $Detect @('-ManifestPath', $Manifest); $la = & $live
    Record 'R2' 'live detection twice: task, HKCU and LocalAppData\PPTK unchanged' ($lb.Task -eq $la.Task -and $lb.Pptk -eq $la.Pptk -and $lb.Reg -eq $la.Reg -and $lb.Local -eq $la.Local -and $l1.Code -ne 0 -and $l1.Out -eq '' -and $l1.Err -eq $l2.Err) (($l1.Err -split "`r?`n")[0])

    # S: Arabic and Unicode folder paths
    $arabic = -join (@(0x0627, 0x062E, 0x062A, 0x0628, 0x0627, 0x0631, 0x0020, 0x0645, 0x0633, 0x062A, 0x062E, 0x062F, 0x0645, 0x0020, 0x2713) | ForEach-Object { [char]$_ })
    $r = Join-Path $Base $arabic; New-Install $r $PkgDir
    Expect-Detected 'S1' 'valid installation under an Arabic/Unicode path' (Run-Root $r)
    Remove-Item -LiteralPath (Join-Path $r "$kd\PPTKKeys.dll")
    Expect-Not 'S2' 'missing file under an Arabic/Unicode path' (Run-Root $r) 'MissingFiles'

    # T: no hard-coded user-specific paths; live mode resolves the current user's folders
    $code = @([IO.File]::ReadAllLines($Detect) | Where-Object { $_ -notmatch '^\s*#' })
    $hits = @($code | Where-Object { $_ -match '(?i)c:\\users|\\appdata\\(roaming|local)|noufal|naldakheel|\$env:USERNAME|\$env:USERPROFILE' })
    Record 'T1' 'no hard-coded user names or profile paths in detect.ps1' ($hits.Count -eq 0) ($hits -join ' / ')
    $who = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    Record 'T2' 'live mode names the current user (resolved at run time)' ($l1.Err -match [regex]::Escape($who)) ''
    $r = New-Case; Edit-Json (Join-Path $r 'identity.json') { param($o) $o.name = 'OTHER\person'; $o.sid = 'S-1-5-21-9' }
    Expect-Not 'T3' 'test identity differs from task user: not detected (no fallback to the real user)' (Run-Root $r) 'IncorrectTask' 'principal'

    # Z: the generated detection script of a real package (embedded version.json)
    if ($Package) {
        $gen = Join-Path (Split-Path -Parent $Package) 'PPTK-detect.ps1'
        $r = Join-Path $Base 'realpkg'; New-Install $r $Package
        Expect-Detected 'Z1' 'generated PPTK-detect.ps1 (embedded manifest) detects a real 1.0.0 install' (Invoke-Detect $gen @('-Root', $r))
        [IO.File]::AppendAllText((Join-Path $r 'AppData\Microsoft\AddIns\PPTK.ppam'), 'x')
        Expect-Not 'Z2' 'generated script: tampered real add-in' (Invoke-Detect $gen @('-Root', $r)) 'HashMismatch'
        $lz = Invoke-Detect $gen @()
        Expect-Not 'Z3' 'generated script, live, no parameters (as Intune runs it)' $lz 'NotInstalled'
    }
} catch {
    Record 'ERR' 'test run' $false $_.Exception.Message
} finally {
    Remove-Item -LiteralPath $Base -Recurse -Force -ErrorAction SilentlyContinue
}

$failed = @($results | Where-Object { -not $_.Ok })
Write-Host ("TOTAL {0} tests, {1} failed" -f $results.Count, $failed.Count)
if ($failed.Count) { exit 1 } else { exit 0 }
