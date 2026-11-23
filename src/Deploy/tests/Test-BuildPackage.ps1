<#
.SYNOPSIS
  Tests for src\Deploy\build-package.ps1.

.DESCRIPTION
  Cases that change the source run on a copy of the needed source files in a temporary folder;
  the repository is never modified. Cases that change a package run on a copy of a built package.
  Builds that reach PowerPoint need PowerPoint to be closed (the build refuses otherwise).

    A  successful build                      F  banned EXE / unlisted file in the package
    B  missing required source               G  hash verification (tampered file, clean package)
    C  missing dependency file               H  rebuild from identical source (reproducibility)
    D  mismatched versions                   I  repeated build leaves no old files; refuses foreign folder
    E  stale or missing add-in version       J  package structure and file inventory

  Exit code 0 when every case passes, 1 otherwise.
#>
param([string]$RepoRoot = '')

$ErrorActionPreference = 'Stop'
if (-not $RepoRoot) { $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path }
$Build = Join-Path $RepoRoot 'src\Deploy\build-package.ps1'
# The release version in the source (PPTKKeys.csproj), and the next patch version for the stale-source case.
$V = ([xml][IO.File]::ReadAllText((Join-Path $RepoRoot 'src\PPTKKeys\windows\PPTKKeys.csproj'))).Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
$Vnext = ([version]$V).Major.ToString() + '.' + ([version]$V).Minor + '.' + (([version]$V).Build + 1)
$TestRoot = Join-Path ([IO.Path]::GetTempPath()) ('PPTK-package-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $TestRoot | Out-Null
$results = New-Object System.Collections.ArrayList

function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) {
    [void]$results.Add([pscustomobject]@{ Id = $id; Name = $name; Ok = $ok; Detail = $detail })
    Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail)
}

function Invoke-Build([string[]]$arguments) {
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File $Build @arguments 2>&1 | ForEach-Object { "$_" }
    return [pscustomobject]@{ Code = $LASTEXITCODE; Text = ($out -join "`n") }
}

function New-SourceCopy([string]$name) {
    $dst = Join-Path $TestRoot $name
    foreach ($rel in 'src\Modules', 'src\Classes', 'src\Forms', 'src\PPTK\Ribbon') {
        Copy-Item -LiteralPath (Join-Path $RepoRoot $rel) -Destination (Join-Path $dst $rel) -Recurse
    }
    New-Item -ItemType Directory -Force -Path (Join-Path $dst 'src\PPTK\Brand') | Out-Null
    Copy-Item -Path (Join-Path $RepoRoot 'src\PPTK\Brand\*.bas'), (Join-Path $RepoRoot 'src\PPTK\Brand\customUI.xml.rels') -Destination (Join-Path $dst 'src\PPTK\Brand')
    Copy-Item -LiteralPath (Join-Path $RepoRoot 'src\PPTK-WINDOWS-FINAL.pptm') -Destination (Join-Path $dst 'src')
    New-Item -ItemType Directory -Force -Path (Join-Path $dst 'src\Deploy') | Out-Null
    foreach ($s in 'detect.ps1', 'install.ps1', 'uninstall.ps1', 'deploy-env.ps1') { Copy-Item -LiteralPath (Join-Path $RepoRoot "src\Deploy\$s") -Destination (Join-Path $dst 'src\Deploy') }
    $keys = Join-Path $dst 'src\PPTKKeys\windows'
    New-Item -ItemType Directory -Force -Path $keys | Out-Null
    Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'src\PPTKKeys\windows') -File | Copy-Item -Destination $keys
    return $dst
}

function Copy-Package([string]$from, [string]$name) {
    $dst = Join-Path $TestRoot $name
    Copy-Item -LiteralPath $from -Destination $dst -Recurse
    return $dst
}

function Get-Inventory([string]$pkg) {
    @(Get-ChildItem -LiteralPath $pkg -Recurse -File | ForEach-Object { $_.FullName.Substring($pkg.Length + 1).Replace('\', '/') } | Sort-Object)
}

try {
    if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Close PowerPoint before running these tests.' }

    # ---------------------------------------------------------------- A, J
    $outA = Join-Path $TestRoot 'outA'
    $a = Invoke-Build @('-RepoRoot', $RepoRoot, '-OutputDir', $outA, '-ExpectedVersion', $V)
    $pkgA = Join-Path $outA 'PPTK-Package'
    Record 'A' 'successful build' ($a.Code -eq 0 -and $a.Text -match ('PACKAGE OK: PPTK ' + [regex]::Escape($V))) "exit $($a.Code)"
    $expected = @('deploy-env.ps1', 'detect.ps1', 'install.ps1', 'payload/PPTK.ppam', 'payload/PPTKKeys/PPTKKeys.deps.json', 'payload/PPTKKeys/PPTKKeys.dll', 'payload/PPTKKeys/PPTKKeys.runtimeconfig.json', 'uninstall.ps1', 'version.json')
    $inv = Get-Inventory $pkgA
    Record 'J' 'package structure and inventory' (($inv -join '|') -eq ($expected -join '|')) ($inv -join ', ')
    $manifestA = Get-Content -LiteralPath (Join-Path $pkgA 'version.json') -Raw | ConvertFrom-Json

    # ---------------------------------------------------------------- B, D, E (fail before PowerPoint)
    $b = New-SourceCopy 'srcB'; Remove-Item -LiteralPath (Join-Path $b 'src\PPTK\Brand\ModulePPTKBrand.bas')
    $r = Invoke-Build @('-RepoRoot', $b, '-OutputDir', (Join-Path $TestRoot 'outB'), '-AllowNoGit')
    Record 'B' 'missing required source' ($r.Code -eq 1 -and $r.Text -match 'required source files are missing: src\\PPTK\\Brand\\ModulePPTKBrand\.bas') (($r.Text -split "`n" | Where-Object { $_ -match 'FAILED' }) -join ' ')

    $d = New-SourceCopy 'srcD'; $csproj = Join-Path $d 'src\PPTKKeys\windows\PPTKKeys.csproj'
    [IO.File]::WriteAllText($csproj, ([IO.File]::ReadAllText($csproj).Replace("<Version>$V</Version>", "<Version>$Vnext</Version>")))
    $r = Invoke-Build @('-RepoRoot', $d, '-OutputDir', (Join-Path $TestRoot 'outD'), '-AllowNoGit')
    Record 'D' 'mismatched versions' ($r.Code -eq 1 -and $r.Text -match 'versions disagree' -and $r.Text -match ('ModulePPTKBrand\.bas\) = ' + [regex]::Escape($V)) -and $r.Text -match ('PPTKKeys\.csproj\) = ' + [regex]::Escape($Vnext))) (($r.Text -split "`n" | Where-Object { $_ -match 'disagree|= 1\.0' }) -join ' / ')

    $r = Invoke-Build @('-RepoRoot', $RepoRoot, '-OutputDir', (Join-Path $TestRoot 'outD2'), '-ExpectedVersion', '2.0.0')
    Record 'D2' 'expected version differs from source' ($r.Code -eq 1 -and $r.Text -match ('expected version 2\.0\.0, but the source says ' + [regex]::Escape($V))) ''

    $e = New-SourceCopy 'srcE1'
    foreach ($f in (Join-Path $e 'src\PPTK\Brand\ModulePPTKBrand.bas'), (Join-Path $e 'src\PPTKKeys\windows\PPTKKeys.csproj')) {
        [IO.File]::WriteAllText($f, ([IO.File]::ReadAllText($f).Replace("`"$V`"", "`"$Vnext`"").Replace("<Version>$V</Version>", "<Version>$Vnext</Version>")))
    }
    $r = Invoke-Build @('-RepoRoot', $e, '-OutputDir', (Join-Path $TestRoot 'outE1'), '-AllowNoGit')
    Record 'E1' "source says $Vnext but the PPTM still has $V (stale add-in source)" ($r.Code -eq 1 -and $r.Text -match 'the PPTM is stale' -and $r.Text -match 'ModulePPTKBrand') (($r.Text -split "`n" | Where-Object { $_ -match 'stale|<>' }) -join ' / ')

    $e2 = New-SourceCopy 'srcE2'; $bas = Join-Path $e2 'src\PPTK\Brand\ModulePPTKBrand.bas'
    [IO.File]::WriteAllLines($bas, @([IO.File]::ReadAllLines($bas) | Where-Object { $_ -notmatch 'PPTK_VERSION As String' }))
    $r = Invoke-Build @('-RepoRoot', $e2, '-OutputDir', (Join-Path $TestRoot 'outE2'), '-AllowNoGit')
    Record 'E2' 'missing PPTK_VERSION' ($r.Code -eq 1 -and $r.Text -match 'PPTK_VERSION \(src\\PPTK\\Brand\\ModulePPTKBrand\.bas\) is missing') ''

    # E3: a package whose add-in has no PPTK_VERSION (an older add-in), with its hash and size
    # updated in version.json so that only the version check can catch it.
    $e3 = Copy-Package $pkgA 'pkgE3'
    # Prefer a real pre-1.0.0 add-in (no PPTK_VERSION) from local backups; else an old PPTM.
    $oldAddin = @(Get-ChildItem -Path (Join-Path $env:USERPROFILE 'PPTK-Backups\*\PPTK.ppam') -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | Select-Object -First 1 | ForEach-Object { $_.FullName })
    $oldAddin = if ($oldAddin.Count) { $oldAddin[0] } else { Join-Path $RepoRoot 'src\PPTK-WINDOWS-FINAL.before-ui-20261105.pptm' }
    if (Test-Path -LiteralPath $oldAddin) {
        Copy-Item -LiteralPath $oldAddin -Destination (Join-Path $e3 'payload\PPTK.ppam') -Force
        $mf = Join-Path $e3 'version.json'; $m = Get-Content -LiteralPath $mf -Raw | ConvertFrom-Json
        $entry = @($m.files | Where-Object { $_.path -eq 'payload/PPTK.ppam' })[0]
        $entry.sha256 = (Get-FileHash -LiteralPath (Join-Path $e3 'payload\PPTK.ppam') -Algorithm SHA256).Hash.ToLowerInvariant()
        $entry.size = (Get-Item -LiteralPath (Join-Path $e3 'payload\PPTK.ppam')).Length
        [IO.File]::WriteAllText($mf, ($m | ConvertTo-Json -Depth 6))
        $r = Invoke-Build @('-VerifyOnly', '-PackagePath', $e3)
        Record 'E3' 'package add-in without the right PPTK_VERSION' ($r.Code -eq 1 -and $r.Text -match "PPTK_VERSION ''|not a macro-enabled PowerPoint add-in") (($r.Text -split "`n" | Where-Object { $_ -match 'FAILED' }) -join ' ')
    } else { Record 'E3' 'package add-in without the right PPTK_VERSION' $false 'old PPTM sample not found' }

    # ---------------------------------------------------------------- C, F, G (verification of package copies)
    $r = Invoke-Build @('-VerifyOnly', '-PackagePath', $pkgA)
    Record 'G1' 'clean package verifies' ($r.Code -eq 0 -and $r.Text -match 'PACKAGE OK') ''

    $c = Copy-Package $pkgA 'pkgC'; Remove-Item -LiteralPath (Join-Path $c 'payload\PPTKKeys\PPTKKeys.runtimeconfig.json')
    $r = Invoke-Build @('-VerifyOnly', '-PackagePath', $c)
    Record 'C' 'missing dependency file' ($r.Code -eq 1 -and $r.Text -match 'listed in version\.json are missing: payload/PPTKKeys/PPTKKeys\.runtimeconfig\.json') ''

    $f = Copy-Package $pkgA 'pkgF'; [IO.File]::WriteAllText((Join-Path $f 'payload\PPTKKeys\PPTKKeys.exe'), 'not a program')
    $r = Invoke-Build @('-VerifyOnly', '-PackagePath', $f)
    Record 'F1' 'banned EXE in the package' ($r.Code -eq 1 -and $r.Text -match 'banned files in the package: payload/PPTKKeys/PPTKKeys\.exe') ''
    $f2 = Copy-Package $pkgA 'pkgF2'; [IO.File]::WriteAllText((Join-Path $f2 'payload\notes.txt'), 'x')
    $r = Invoke-Build @('-VerifyOnly', '-PackagePath', $f2)
    Record 'F2' 'unlisted file in the package' ($r.Code -eq 1 -and $r.Text -match 'does not list: payload/notes\.txt') ''
    $f3 = Copy-Package $pkgA 'pkgF3'; [IO.File]::WriteAllText((Join-Path $f3 'payload\PPTKKeys\PPTKKeys.pdb'), 'x')
    $r = Invoke-Build @('-VerifyOnly', '-PackagePath', $f3)
    Record 'F3' 'debug symbols in the package' ($r.Code -eq 1 -and $r.Text -match 'banned files.*PPTKKeys\.pdb') ''

    $g = Copy-Package $pkgA 'pkgG'; $dep = Join-Path $g 'payload\PPTKKeys\PPTKKeys.deps.json'
    [IO.File]::WriteAllText($dep, ([IO.File]::ReadAllText($dep) + ' '))
    $r = Invoke-Build @('-VerifyOnly', '-PackagePath', $g)
    Record 'G2' 'tampered file is detected' ($r.Code -eq 1 -and $r.Text -match 'SHA-256 mismatch for payload/PPTKKeys/PPTKKeys\.deps\.json') ''
    $g3 = Copy-Package $pkgA 'pkgG3'; $mf = Join-Path $g3 'version.json'; $m = Get-Content -LiteralPath $mf -Raw | ConvertFrom-Json
    @($m.files | Where-Object { $_.path -eq 'payload/PPTKKeys/PPTKKeys.dll' })[0].sha256 = ('0' * 64)
    [IO.File]::WriteAllText($mf, ($m | ConvertTo-Json -Depth 6))
    $r = Invoke-Build @('-VerifyOnly', '-PackagePath', $g3)
    Record 'G3' 'wrong hash in version.json is detected' ($r.Code -eq 1 -and $r.Text -match 'SHA-256 mismatch for payload/PPTKKeys/PPTKKeys\.dll') ''
    $okHashes = $true
    foreach ($fe in $manifestA.files) { if ((Get-FileHash -LiteralPath (Join-Path $pkgA ($fe.path.Replace('/', '\'))) -Algorithm SHA256).Hash.ToLowerInvariant() -ne $fe.sha256) { $okHashes = $false } }
    Record 'G4' 'every SHA-256 in version.json matches the file (independent check)' $okHashes ''

    # ---------------------------------------------------------------- H (second build, same source)
    $outH = Join-Path $TestRoot 'outH'
    $h = Invoke-Build @('-RepoRoot', $RepoRoot, '-OutputDir', $outH, '-ExpectedVersion', $V)
    $manifestH = Get-Content -LiteralPath (Join-Path $outH 'PPTK-Package\version.json') -Raw | ConvertFrom-Json
    $detSame = $true; $lines = @()
    foreach ($fa in $manifestA.files) {
        $fh = @($manifestH.files | Where-Object { $_.path -eq $fa.path })[0]
        if ($fa.deterministic -and $fa.sha256 -ne $fh.sha256) { $detSame = $false }
        $lines += "{0}: {1} {2}" -f $fa.path, $(if ($fa.sha256 -eq $fh.sha256) { 'same bytes' } else { 'different bytes' }), $(if ($fa.contentFingerprint) { 'fingerprint ' + $(if ($fa.contentFingerprint -eq $fh.contentFingerprint) { 'same' } else { 'DIFFERENT' }) })
    }
    $fpSame = (@($manifestA.files | Where-Object { $_.path -eq 'payload/PPTK.ppam' })[0].contentFingerprint -eq @($manifestH.files | Where-Object { $_.path -eq 'payload/PPTK.ppam' })[0].contentFingerprint)
    Record 'H' 'rebuild from identical source' ($h.Code -eq 0 -and $detSame -and $fpSame) ($lines -join ' ; ')

    # ---------------------------------------------------------------- I (repeated build into the same folder)
    [IO.File]::WriteAllText((Join-Path $pkgA 'payload\leftover.tmp'), 'old')
    $i1 = Invoke-Build @('-RepoRoot', $RepoRoot, '-OutputDir', $outA, '-ExpectedVersion', $V)
    $inv2 = Get-Inventory $pkgA
    Record 'I1' 'repeated build leaves no old files' ($i1.Code -eq 0 -and ($inv2 -join '|') -eq ($expected -join '|')) ($inv2 -join ', ')
    $foreign = Join-Path $TestRoot 'outForeign\PPTK-Package'; New-Item -ItemType Directory -Force -Path $foreign | Out-Null; [IO.File]::WriteAllText((Join-Path $foreign 'keep.txt'), 'not a package')
    $i2 = Invoke-Build @('-RepoRoot', $RepoRoot, '-OutputDir', (Join-Path $TestRoot 'outForeign'), '-ExpectedVersion', $V)
    Record 'I2' 'refuses to delete a folder that is not a PPTK package' ($i2.Code -eq 1 -and $i2.Text -match 'is not a PPTK package; not deleting it' -and (Test-Path -LiteralPath (Join-Path $foreign 'keep.txt'))) ''

    Record 'X' 'no PowerPoint or temp build folder left behind' (-not (Get-Process POWERPNT -ErrorAction SilentlyContinue) -and @(Get-ChildItem ([IO.Path]::GetTempPath()) -Directory -Filter 'PPTK-package-build-*').Count -eq 0) ''
} catch {
    Record 'ERR' 'test run' $false $_.Exception.Message
} finally {
    Remove-Item -LiteralPath $TestRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$failed = @($results | Where-Object { -not $_.Ok })
Write-Host ("TOTAL {0} tests, {1} failed" -f $results.Count, $failed.Count)
if ($failed.Count) { exit 1 } else { exit 0 }
