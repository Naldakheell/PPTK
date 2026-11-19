<#
.SYNOPSIS
  Builds the versioned PPTK deployment package from the project source.

.DESCRIPTION
  Output (default <repo>\dist\PPTK-Package):

    PPTK-Package\
      version.json
      payload\PPTK.ppam
      payload\PPTKKeys\PPTKKeys.dll
      payload\PPTKKeys\PPTKKeys.deps.json
      payload\PPTKKeys\PPTKKeys.runtimeconfig.json

  Sources (never the installed add-in, the installed helper, release\win-x64, or backups):
    PPTK.ppam  <- src\PPTK-WINDOWS-FINAL.pptm, saved as an add-in by PowerPoint (ppSaveAsOpenXMLAddin).
                  The VBA source stored in the PPTM is read from the file (MS-OVBA) and compared with
                  the source files in src (Modules, Classes, Forms, PPTK\Brand) before PowerPoint
                  starts; a stale PPTM stops the build. The VBA in the built add-in is read the same
                  way and must equal the PPTM's.
    PPTKKeys   <- dotnet publish of src\PPTKKeys\windows\PPTKKeys.csproj with UseAppHost=false,
                  so no PPTKKeys.exe is produced. Only the DLL and its two JSON files are packaged.

  Versions: PPTK_VERSION (src\PPTK\Brand\ModulePPTKBrand.bas) and <Version> (PPTKKeys.csproj)
  must be equal. The built add-in must contain that PPTK_VERSION (read from the add-in file) and
  PPTK_GetVersion() must return it; the DLL must carry it (assembly version and --version).
  Any mismatch stops the build and names the components.

  Reproducibility: PPTKKeys files are byte-identical for the same source. PowerPoint rewrites
  ppt/vbaProject.bin on every save (compiled-code caches and project protection values), so the
  PPAM bytes differ between builds. version.json therefore also records contentFingerprint: a
  SHA-256 over every other part of the add-in and the VBA source of every component, read from
  the add-in file. It is identical for the same source.

  Build machine: Windows, desktop PowerPoint, .NET 10 SDK, Git. PowerPoint must not be running:
  the build never uses a PowerPoint session that is in use, and it closes only the PowerPoint
  instance it started.

  Next to the package it also writes PPTK-detect.ps1: src\Deploy\detect.ps1 with this package's
  version.json embedded, for upload as the Intune custom detection script (Intune runs it without
  parameters and without the package content).

  The deployment scripts listed in $PackageScripts (install, uninstall, detect, deploy-env) are
  (copied from src\Deploy\scripts to the package root and hashed like every other file).

.PARAMETER RuntimeX64 / RuntimeArm64
  Universal package: the extracted official Microsoft .NET runtime archives (dotnet-runtime-<v>-win-x64.zip and
  dotnet-runtime-<v>-win-arm64.zip, verified against Microsoft's release metadata before the build). Both are
  copied to payload\runtime\<rid>\ and listed file by file in version.json, with a fingerprint per runtime in
  "runtimes". The build checks that every .exe/.dll is validly signed by Microsoft, that each dotnet.exe is built
  for its architecture, and that each runtime has the framework PPTKKeys.runtimeconfig.json asks for.
  install.ps1 stages only the runtime of the device's architecture; the task then runs that private dotnet.exe.
  PPTK-detect.ps1 embeds a slim manifest (no runtime file list, the fingerprints instead).

.PARAMETER VerifyOnly
  Validates an existing package (inventory, hashes, banned files, add-in structure and version,
  DLL version, content fingerprint) without building. Needs no PowerPoint.

.NOTES
  Exit code 0 on success, 1 on any error. Writes only the output folder and the build's own
  temporary folder (dotnet also writes its usual Git-ignored bin\obj folders).
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = '',
    [string]$OutputDir = '',
    [string]$ExpectedVersion = '',
    [string]$DotnetPath = '',
    [switch]$VerifyOnly,
    [string]$PackagePath = '',
    # Tests only: build from a source copy that is not a Git working tree.
    [switch]$AllowNoGit,
    [string]$RuntimeX64 = '',
    [string]$RuntimeArm64 = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

# Deployment scripts copied from srcDeploy to the package root (listed and hashed in version.json).
$PackageScripts = @('install.ps1', 'uninstall.ps1', 'detect.ps1', 'deploy-env.ps1')

$PackageName = 'PPTK'
$PackageFolderName = 'PPTK-Package'
$AddinContentType = 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml'
$KeysFiles = @('PPTKKeys.dll', 'PPTKKeys.deps.json', 'PPTKKeys.runtimeconfig.json')
$VersionPattern = '(?m)^\s*Public Const PPTK_VERSION As String = "([^"]*)"'

function Fail([string]$message) { throw "BUILD FAILED: $message" }
function Note([string]$message) { Write-Host "  $message" }
function Step([string]$message) { Write-Host "== $message" }

# Reads the VBA source of every module from a vbaProject.bin (compound file + MS-OVBA
# compression). Managed code only; used for the source check, the version check and the
# content fingerprint, so none of them needs PowerPoint or access to the VBA object model.
Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public static class PptkVbaReader
{
    sealed class Entry { public string Name; public int Type; public uint Left, Right, Child, Start; public long Size; }

    sealed class Cfb
    {
        readonly byte[] d; readonly int ss, mss; readonly uint cutoff;
        readonly List<uint> fat = new List<uint>(); readonly List<uint> minifat = new List<uint>();
        public readonly List<Entry> Entries = new List<Entry>(); readonly byte[] mini;

        public Cfb(byte[] data)
        {
            d = data;
            byte[] sig = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            for (int i = 0; i < 8; i++) if (d[i] != sig[i]) throw new InvalidDataException("vbaProject.bin is not a compound file");
            ss = 1 << BitConverter.ToUInt16(d, 0x1E); mss = 1 << BitConverter.ToUInt16(d, 0x20);
            uint nfat = BitConverter.ToUInt32(d, 0x2C), firstDir = BitConverter.ToUInt32(d, 0x30);
            cutoff = BitConverter.ToUInt32(d, 0x38);
            uint firstMini = BitConverter.ToUInt32(d, 0x3C), sid = BitConverter.ToUInt32(d, 0x44);
            var difat = new List<uint>();
            for (int i = 0; i < 109; i++) difat.Add(BitConverter.ToUInt32(d, 0x4C + 4 * i));
            int guard = 0;
            while (sid < 0xFFFFFFFA && guard++ < 100000)
            {
                int off = (int)((sid + 1) * ss), n = ss / 4;
                for (int i = 0; i < n - 1; i++) difat.Add(BitConverter.ToUInt32(d, off + 4 * i));
                sid = BitConverter.ToUInt32(d, off + 4 * (n - 1));
            }
            for (int k = 0; k < nfat; k++)
            {
                int off = (int)((difat[k] + 1) * ss);
                for (int i = 0; i < ss / 4; i++) fat.Add(BitConverter.ToUInt32(d, off + 4 * i));
            }
            byte[] dir = Chain(firstDir);
            for (int i = 0; i + 128 <= dir.Length; i += 128)
            {
                int nlen = BitConverter.ToUInt16(dir, i + 0x40);
                var e = new Entry();
                e.Name = nlen >= 2 ? Encoding.Unicode.GetString(dir, i, nlen - 2) : "";
                e.Type = dir[i + 0x42];
                e.Left = BitConverter.ToUInt32(dir, i + 0x44); e.Right = BitConverter.ToUInt32(dir, i + 0x48); e.Child = BitConverter.ToUInt32(dir, i + 0x4C);
                e.Start = BitConverter.ToUInt32(dir, i + 0x74); e.Size = (long)(BitConverter.ToUInt64(dir, i + 0x78) & 0xFFFFFFFF);
                Entries.Add(e);
            }
            byte[] rootData = Chain(Entries[0].Start);
            mini = new byte[Math.Min(rootData.Length, Entries[0].Size)];
            Array.Copy(rootData, mini, mini.Length);
            if (firstMini < 0xFFFFFFFA) { byte[] mf = Chain(firstMini); for (int i = 0; i + 4 <= mf.Length; i += 4) minifat.Add(BitConverter.ToUInt32(mf, i)); }
        }

        byte[] Chain(uint sid)
        {
            var ms = new MemoryStream(); int guard = 0;
            while (sid < 0xFFFFFFFA && guard++ < 1000000) { ms.Write(d, (int)((sid + 1) * ss), ss); sid = fat[(int)sid]; }
            return ms.ToArray();
        }

        public byte[] Read(Entry e)
        {
            var ms = new MemoryStream();
            if (e.Size < cutoff)
            {
                uint sid = e.Start; int guard = 0;
                while (sid < 0xFFFFFFFA && guard++ < 1000000) { ms.Write(mini, (int)(sid * mss), mss); sid = minifat[(int)sid]; }
            }
            else { byte[] c = Chain(e.Start); ms.Write(c, 0, c.Length); }
            byte[] all = ms.ToArray(); byte[] r = new byte[Math.Min(all.Length, e.Size)]; Array.Copy(all, r, r.Length); return r;
        }

        public int Find(int parent, string name)
        {
            var stack = new Stack<uint>(); stack.Push(Entries[parent].Child);
            while (stack.Count > 0)
            {
                uint i = stack.Pop();
                if (i >= 0xFFFFFFFA || i >= Entries.Count) continue;
                if (string.Equals(Entries[(int)i].Name, name, StringComparison.OrdinalIgnoreCase)) return (int)i;
                stack.Push(Entries[(int)i].Left); stack.Push(Entries[(int)i].Right);
            }
            return -1;
        }
    }

    static byte[] Decompress(byte[] buf, int start)
    {
        if (start >= buf.Length || buf[start] != 1) throw new InvalidDataException("bad MS-OVBA compressed container");
        var o = new List<byte>(); int pos = start + 1;
        while (pos + 1 < buf.Length)
        {
            int hdr = BitConverter.ToUInt16(buf, pos); int size = (hdr & 0x0FFF) + 3; bool compressed = (hdr & 0x8000) != 0;
            int chunkEnd = Math.Min(pos + size, buf.Length); pos += 2; int chunkStart = o.Count;
            if (!compressed) { int n = Math.Min(4096, buf.Length - pos); for (int i = 0; i < n; i++) o.Add(buf[pos + i]); pos += 4096; continue; }
            while (pos < chunkEnd)
            {
                byte flags = buf[pos++];
                for (int bit = 0; bit < 8; bit++)
                {
                    if (pos >= chunkEnd) break;
                    if (((flags >> bit) & 1) == 0) { o.Add(buf[pos++]); continue; }
                    int token = BitConverter.ToUInt16(buf, pos); pos += 2;
                    int diff = o.Count - chunkStart; int bc = 4; while ((1 << bc) < diff) bc++;
                    int lmask = 0xFFFF >> bc; int length = (token & lmask) + 3; int offset = (token >> (16 - bc)) + 1;
                    for (int i = 0; i < length; i++) o.Add(o[o.Count - offset]);
                }
            }
        }
        return o.ToArray();
    }

    // Module name -> source text (as stored, including Attribute lines).
    public static Dictionary<string, string> ReadModules(byte[] vbaProjectBin)
    {
        var cfb = new Cfb(vbaProjectBin);
        int vba = cfb.Find(0, "VBA"); if (vba < 0) throw new InvalidDataException("no VBA storage");
        int dirIdx = cfb.Find(vba, "dir"); if (dirIdx < 0) throw new InvalidDataException("no VBA/dir stream");
        byte[] dir = Decompress(cfb.Read(cfb.Entries[dirIdx]), 0);
        int codepage = 1252; var mods = new List<string[]>(); string name = null, stream = null; uint offset = 0;
        int p = 0;
        while (p + 6 <= dir.Length)
        {
            int id = BitConverter.ToUInt16(dir, p); int size = (int)BitConverter.ToUInt32(dir, p + 2); p += 6;
            if (id == 0x0009) { p += 6; continue; }
            if (p + size > dir.Length) break;
            if (id == 0x0003) codepage = BitConverter.ToUInt16(dir, p);
            else if (id == 0x0019) { name = Encoding.GetEncoding(28591).GetString(dir, p, size); stream = null; offset = 0; }
            else if (id == 0x001A && name != null) stream = Encoding.GetEncoding(28591).GetString(dir, p, size);
            else if (id == 0x0031 && name != null) offset = BitConverter.ToUInt32(dir, p);
            else if (id == 0x002B && name != null) { mods.Add(new[] { name, stream, offset.ToString() }); name = null; }
            p += size;
        }
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var enc = Encoding.GetEncoding(codepage);
        foreach (var m in mods)
        {
            int si = cfb.Find(vba, m[1]); if (si < 0) throw new InvalidDataException("missing module stream " + m[1]);
            result[m[0]] = enc.GetString(Decompress(cfb.Read(cfb.Entries[si]), int.Parse(m[2])));
        }
        return result;
    }
}
'@

function Get-Sha256([string]$path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Get-Sha256Bytes([byte[]]$bytes) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')).ToLowerInvariant() } finally { $sha.Dispose() }
}
function Get-Sha256Text([string]$text) { Get-Sha256Bytes ([Text.Encoding]::UTF8.GetBytes($text)) }
function Test-SemVer([string]$value) { return $value -match '^\d+\.\d+\.\d+$' }

# ---------------------------------------------------------------- versions
function Resolve-PackageVersion([string]$root, [string]$expected) {
    $found = [ordered]@{ 'PPTK_VERSION (src\PPTK\Brand\ModulePPTKBrand.bas)' = $null; '<Version> (src\PPTKKeys\windows\PPTKKeys.csproj)' = $null }
    $m = [regex]::Match([IO.File]::ReadAllText((Join-Path $root 'src\PPTK\Brand\ModulePPTKBrand.bas')), $VersionPattern)
    if ($m.Success) { $found['PPTK_VERSION (src\PPTK\Brand\ModulePPTKBrand.bas)'] = $m.Groups[1].Value }
    $m = [regex]::Match([IO.File]::ReadAllText((Join-Path $root 'src\PPTKKeys\windows\PPTKKeys.csproj')), '<Version>\s*([^<]*?)\s*</Version>')
    if ($m.Success) { $found['<Version> (src\PPTKKeys\windows\PPTKKeys.csproj)'] = $m.Groups[1].Value }
    $problems = @()
    foreach ($k in $found.Keys) {
        if (-not $found[$k]) { $problems += "$k is missing" }
        elseif (-not (Test-SemVer $found[$k])) { $problems += "$k = '$($found[$k])' is not MAJOR.MINOR.PATCH" }
    }
    if ($problems.Count) { Fail ("version error:`n    " + ($problems -join "`n    ")) }
    $distinct = @($found.Values | Select-Object -Unique)
    if ($distinct.Count -ne 1) { Fail ("versions disagree:`n    " + (($found.Keys | ForEach-Object { "$_ = $($found[$_])" }) -join "`n    ")) }
    if ($expected -and $expected -ne $distinct[0]) { Fail "expected version $expected, but the source says $($distinct[0])" }
    return $distinct[0]
}

# ---------------------------------------------------------------- VBA from files
function Read-ZipParts([string]$path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $parts = [ordered]@{}
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        foreach ($e in $zip.Entries) {
            $ms = New-Object IO.MemoryStream
            $st = $e.Open(); try { $st.CopyTo($ms) } finally { $st.Dispose() }
            $parts[$e.FullName] = $ms.ToArray()
        }
    } finally { $zip.Dispose() }
    return $parts
}

function Get-VbaModules($parts, [string]$label) {
    if (-not $parts.Contains('ppt/vbaProject.bin')) { Fail "$label has no ppt/vbaProject.bin" }
    try { return [PptkVbaReader]::ReadModules($parts['ppt/vbaProject.bin']) } catch { Fail "cannot read the VBA in ${label}: $($_.Exception.Message)" }
}

# Code lines without the .frm/.cls designer header and without Attribute lines.
function Get-CodeLines([string[]]$lines) {
    if ($lines.Count -gt 0 -and $lines[0].StartsWith('VERSION')) {
        $first = -1
        for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i].StartsWith('Attribute ')) { $first = $i; break } }
        if ($first -ge 0) { $lines = $lines[$first..($lines.Count - 1)] }
    }
    return @($lines | Where-Object { -not $_.StartsWith('Attribute ') })
}
# The VBA editor re-cases identifiers and inserts spaces ("f (x)"); compare without case and whitespace.
function Get-Comparable([string[]]$lines) { (@($lines | ForEach-Object { ($_ -replace '\s', '').ToLowerInvariant() } | Where-Object { $_ }) -join "`n") }

function Find-SourceFile([string]$root, [string]$name) {
    foreach ($c in "src\Modules\$name.bas", "src\PPTK\Brand\$name.bas", "src\Classes\$name.cls", "src\Forms\$name.frm") {
        $p = Join-Path $root $c; if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

function Get-ModuleVersion($modules) {
    if (-not $modules.ContainsKey('ModulePPTKBrand')) { return $null }
    $m = [regex]::Match($modules['ModulePPTKBrand'], $VersionPattern)
    if ($m.Success) { return $m.Groups[1].Value } else { return $null }
}

function Get-ContentFingerprint($parts, $modules) {
    $lines = @($parts.Keys | Where-Object { $_ -ne 'ppt/vbaProject.bin' } | Sort-Object | ForEach-Object { 'part:' + $_ + ':' + (Get-Sha256Bytes $parts[$_]) })
    # Code only: PowerPoint rewrites some form Attribute lines when it saves an add-in.
    $lines += @($modules.Keys | Sort-Object | ForEach-Object { 'vba:' + $_ + ':' + (Get-Sha256Text ((Get-CodeLines ($modules[$_] -split "`r?`n") | ForEach-Object { $_.TrimEnd() }) -join "`n")) })
    return Get-Sha256Text ($lines -join "`n")
}

function Test-PpamStructure($parts, [string]$label, [string]$root) {
    foreach ($req in '[Content_Types].xml', 'ppt/vbaProject.bin', 'customUI/customUI.xml', 'customUI/_rels/customUI.xml.rels', '_rels/.rels') {
        if (-not $parts.Contains($req)) { Fail "$label is missing the part $req" }
    }
    # The ribbon schema and the root relationship that loads it must match, or PowerPoint ignores the whole ribbon:
    # 2009/07 schema (Office 2010+, needed for the right-click Select menu, from 1.2.0) <-> 2007 ui/extensibility
    # relationship; 2006/01 schema (releases up to 1.1.0) <-> 2006 relationship. Older releases stay verifiable;
    # a new build is held to the 2009/07 schema by the comparison with src\PPTK\Ribbon\CustomUI.xml below.
    $rootRels = [Text.Encoding]::UTF8.GetString($parts['_rels/.rels'])
    $uiText = [Text.Encoding]::UTF8.GetString($parts['customUI/customUI.xml'])
    $relYear = $null
    foreach ($year in '2007', '2006') {
        $t = "http://schemas\.microsoft\.com/office/$year/relationships/ui/extensibility"
        if ($rootRels -match "Type=""$t""\s+Target=""customUI/customUI\.xml""" -or $rootRels -match "Target=""customUI/customUI\.xml""[^>]*Type=""$t""") { $relYear = $year; break }
    }
    if (-not $relYear) { Fail "$label does not link customUI/customUI.xml with a ui/extensibility relationship" }
    if ($uiText -match 'xmlns="http://schemas\.microsoft\.com/office/2009/07/customui"') { $uiSchema = '2009/07' }
    elseif ($uiText -match 'xmlns="http://schemas\.microsoft\.com/office/2006/01/customui"') { $uiSchema = '2006/01' }
    else { Fail "$label ribbon has no known customUI schema" }
    if (($uiSchema -eq '2009/07') -ne ($relYear -eq '2007')) { Fail "$label ribbon schema $uiSchema does not match its $relYear ui/extensibility relationship (PowerPoint would not load the ribbon)" }
    if ([Text.Encoding]::UTF8.GetString($parts['[Content_Types].xml']) -notmatch [regex]::Escape($AddinContentType)) { Fail "$label is not a macro-enabled PowerPoint add-in ($AddinContentType not found)" }
    $relsText = [Text.Encoding]::UTF8.GetString($parts['customUI/_rels/customUI.xml.rels'])
    $targets = @([regex]::Matches($relsText, 'Target="images/([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    $missing = @($targets | Where-Object { -not $parts.Contains("customUI/images/$_") })
    if ($missing.Count) { Fail ("$label ribbon images missing: " + ($missing -join ', ')) }
    if ($root) {
        if ((Get-Sha256Bytes $parts['customUI/customUI.xml']) -ne (Get-Sha256 (Join-Path $root 'src\PPTK\Ribbon\CustomUI.xml'))) { Fail "$label ribbon (customUI.xml) differs from src\PPTK\Ribbon\CustomUI.xml" }
        $pairs = { param($t) @([regex]::Matches($t, 'Id="([^"]+)"[^>]*Target="([^"]+)"') | ForEach-Object { $_.Groups[1].Value + '=' + $_.Groups[2].Value } | Sort-Object) -join '|' }
        if ((& $pairs $relsText) -ne (& $pairs ([IO.File]::ReadAllText((Join-Path $root 'src\PPTK\Brand\customUI.xml.rels'))))) { Fail "$label ribbon relationships differ from src\PPTK\Brand\customUI.xml.rels" }
    }
    return $targets.Count
}

# ---------------------------------------------------------------- package verification (also -VerifyOnly)
# ---------------------------------------------------------------- private .NET runtimes (universal package)
$RuntimeMachine = @{ 'win-x64' = 0x8664; 'win-arm64' = 0xAA64 }
function Get-PeMachine([string]$path) {
    $b = [IO.File]::ReadAllBytes($path); return [int][BitConverter]::ToUInt16($b, [BitConverter]::ToInt32($b, 0x3c) + 4)
}
# Same definition as deploy-env.ps1 and detect.ps1.
function Get-RuntimeFingerprint([string]$dir) {
    $dir = (Get-Item -LiteralPath $dir).FullName.TrimEnd('\')   # the long form: Get-ChildItem returns long names under an 8.3 path
    $lines = [Collections.Generic.List[string]]::new()
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Recurse -File -Force)) { $lines.Add($f.FullName.Substring($dir.Length + 1).Replace('\', '/') + ' ' + (Get-Sha256 $f.FullName)) }
    $lines.Sort([StringComparer]::Ordinal)
    return Get-Sha256Bytes ([Text.Encoding]::UTF8.GetBytes(($lines -join "`n") + "`n"))
}
function Test-Runtime([string]$dir, [string]$rid, [string]$framework, [string]$minimum) {
    $exe = Join-Path $dir 'dotnet.exe'
    if (-not (Test-Path -LiteralPath $exe)) { Fail "the $rid runtime has no dotnet.exe ($dir)" }
    if ((Get-PeMachine $exe) -ne $RuntimeMachine[$rid]) { Fail ('the {0} runtime''s dotnet.exe is built for machine 0x{1:X4}' -f $rid, (Get-PeMachine $exe)) }
    $bins = @(Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.Extension -in '.exe', '.dll' })
    $bad = @($bins | Where-Object { $s = Get-AuthenticodeSignature -LiteralPath $_.FullName; $s.Status -ne 'Valid' -or $s.SignerCertificate.Subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)' })
    if ($bad.Count) { Fail ("not validly signed by Microsoft in the $rid runtime: " + (($bad | Select-Object -First 5 | ForEach-Object { $_.Name }) -join ', ')) }
    $min = [version]$minimum
    $ver = @(Get-ChildItem -LiteralPath (Join-Path $dir "shared\$framework") -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d+\.\d+\.\d+$' -and ([version]$_.Name).Major -eq $min.Major -and [version]$_.Name -ge $min } | ForEach-Object { $_.Name })
    if (-not $ver.Count) { Fail "the $rid runtime has no $framework $minimum (same major version)" }
    $odd = @(Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.FullName.Substring($dir.Length + 1).Replace('\', '/') -notmatch '^[A-Za-z0-9._-]+(/[A-Za-z0-9._-]+)*$' })
    if ($odd.Count) { Fail "the $rid runtime has file names the manifest cannot list: $($odd[0].Name)" }
    return [pscustomobject]@{ Version = ($ver | Sort-Object { [version]$_ } | Select-Object -Last 1); Files = $bins.Count; Signed = $bins.Count }
}

function Test-Package([string]$pkg) {
    Step "Verify package $pkg"
    $manifestPath = Join-Path $pkg 'version.json'
    if (-not (Test-Path -LiteralPath $manifestPath)) { Fail 'version.json is missing' }
    try { $m = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { Fail "version.json is not valid JSON: $($_.Exception.Message)" }
    if ($m.packageName -ne $PackageName) { Fail "version.json packageName is '$($m.packageName)', expected '$PackageName'" }
    if (-not (Test-SemVer $m.packageVersion)) { Fail "version.json packageVersion '$($m.packageVersion)' is not MAJOR.MINOR.PATCH" }

    $all = @(Get-ChildItem -LiteralPath $pkg -Recurse -File -Force | ForEach-Object { $_.FullName.Substring($pkg.Length + 1).Replace('\', '/') })
    # Executables only inside the official runtimes (payload/runtime/<rid>/, signature-checked below).
    $banned = @($all | Where-Object { ($_ -match '\.(exe|pdb|msi|dll\.config|bat|cmd|com|scr|vbs|js)$' -and $_ -notmatch '^payload/runtime/(win-x64|win-arm64)/') -or $_ -match '(^|/)(bin|obj|release)/' -or ($_ -match '^payload/runtime/' -and $_ -notmatch '^payload/runtime/(win-x64|win-arm64)/') })
    if ($banned.Count) { Fail ("banned files in the package: " + ($banned -join ', ')) }

    $listed = @($m.files | ForEach-Object { $_.path })
    $required = @('payload/PPTK.ppam') + @($KeysFiles | ForEach-Object { "payload/PPTKKeys/$_" }) + @($PackageScripts)
    $missingReq = @($required | Where-Object { $listed -notcontains $_ })
    if ($missingReq.Count) { Fail ("version.json does not list required files: " + ($missingReq -join ', ')) }
    $notOnDisk = @($listed | Where-Object { $all -notcontains $_ })
    if ($notOnDisk.Count) { Fail ("files listed in version.json are missing: " + ($notOnDisk -join ', ')) }
    $extra = @($all | Where-Object { $_ -ne 'version.json' -and $listed -notcontains $_ })
    if ($extra.Count) { Fail ("files in the package that version.json does not list: " + ($extra -join ', ')) }

    foreach ($f in $m.files) {
        $p = Join-Path $pkg ($f.path.Replace('/', '\'))
        $h = Get-Sha256 $p
        if ($h -ne $f.sha256) { Fail "SHA-256 mismatch for $($f.path): file $h, version.json $($f.sha256)" }
        if ((Get-Item -LiteralPath $p).Length -ne [long]$f.size) { Fail "size mismatch for $($f.path)" }
    }

    $dll = Join-Path $pkg 'payload\PPTKKeys\PPTKKeys.dll'
    $asmVersion = [Reflection.AssemblyName]::GetAssemblyName($dll).Version.ToString(3)
    if ($asmVersion -ne $m.packageVersion) { Fail "PPTKKeys.dll version is $asmVersion, package version is $($m.packageVersion)" }
    $rc = Get-Content -LiteralPath (Join-Path $pkg 'payload\PPTKKeys\PPTKKeys.runtimeconfig.json') -Raw | ConvertFrom-Json
    if ($rc.runtimeOptions.framework.name -ne $m.requirements.dotnetRuntime.name -or $rc.runtimeOptions.framework.version -ne $m.requirements.dotnetRuntime.minimumVersion) {
        Fail 'PPTKKeys.runtimeconfig.json does not match requirements.dotnetRuntime in version.json'
    }

    $rids = @($all | ForEach-Object { if ($_ -match '^payload/runtime/([^/]+)/') { $Matches[1] } } | Select-Object -Unique)
    foreach ($rid in $rids) {
        $dir = Join-Path $pkg "payload\runtime\$rid"
        $rt = Test-Runtime $dir $rid $rc.runtimeOptions.framework.name $rc.runtimeOptions.framework.version
        if (-not ($m.PSObject.Properties.Name -contains 'runtimes') -or (Get-RuntimeFingerprint $dir) -ne $m.runtimes.$rid.fingerprint) { Fail "the $rid runtime does not match its fingerprint in version.json" }
        Note "runtime $rid`: .NET $($rt.Version), $($rt.Signed) binaries signed by Microsoft, dotnet.exe for $rid, fingerprint matches"
    }
    if ($rids.Count -and (@($rids | Sort-Object) -join ',') -ne 'win-arm64,win-x64') { Fail "a universal package needs both win-x64 and win-arm64 runtimes, found: $($rids -join ', ')" }

    $ppamEntry = @($m.files | Where-Object { $_.path -eq 'payload/PPTK.ppam' })[0]
    $parts = Read-ZipParts (Join-Path $pkg 'payload\PPTK.ppam')
    $images = Test-PpamStructure $parts 'PPTK.ppam' ''
    $modules = Get-VbaModules $parts 'PPTK.ppam'
    $ppamVersion = Get-ModuleVersion $modules
    if ($ppamVersion -ne $m.packageVersion) { Fail "PPTK.ppam contains PPTK_VERSION '$ppamVersion', package version is $($m.packageVersion)" }
    $fp = Get-ContentFingerprint $parts $modules
    if ($fp -ne $ppamEntry.contentFingerprint) { Fail "PPTK.ppam content fingerprint $fp does not match version.json" }
    Note "files: $($all.Count) (version.json + $($listed.Count) listed); SHA-256 and sizes match; no banned files"
    Note "PPTKKeys.dll $asmVersion; PPTK.ppam PPTK_VERSION $ppamVersion, $($modules.Count) VBA modules, $images ribbon images, fingerprint matches"
    return $m
}

# ---------------------------------------------------------------- main
try {
    if ($VerifyOnly) {
        if (-not $PackagePath) { Fail '-VerifyOnly needs -PackagePath' }
        $pkgFull = (Resolve-Path -LiteralPath $PackagePath).Path.TrimEnd('\')
        $m = Test-Package $pkgFull
        Write-Host "PACKAGE OK: $($m.packageName) $($m.packageVersion)"
        exit 0
    }

    if (-not $RepoRoot) { $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path }
    $RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path.TrimEnd('\')
    if (-not $OutputDir) { $OutputDir = Join-Path $RepoRoot 'dist' }
    $Runtimes = [ordered]@{}
    if ($RuntimeX64 -or $RuntimeArm64) {
        if (-not ($RuntimeX64 -and $RuntimeArm64)) { Fail 'a universal package needs both -RuntimeX64 and -RuntimeArm64' }
        foreach ($pair in @(@('win-x64', $RuntimeX64), @('win-arm64', $RuntimeArm64))) {
            if (-not (Test-Path -LiteralPath (Join-Path $pair[1] 'dotnet.exe'))) { Fail "$($pair[0]) runtime folder has no dotnet.exe: $($pair[1])" }
            $Runtimes[$pair[0]] = (Resolve-Path -LiteralPath $pair[1]).Path.TrimEnd('\')
        }
    }

    Step 'Check sources'
    $requiredSources = @(
        'src\PPTK-WINDOWS-FINAL.pptm',
        'src\PPTK\Brand\ModulePPTKBrand.bas',
        'src\PPTK\Ribbon\CustomUI.xml',
        'src\PPTK\Brand\customUI.xml.rels',
        'src\PPTKKeys\windows\PPTKKeys.csproj',
        'src\PPTKKeys\windows\Program.cs',
        'src\Deploy\detect.ps1'
    ) + @($PackageScripts | ForEach-Object { "src\Deploy\$_" })
    $missingSrc = @($requiredSources | Where-Object { -not (Test-Path -LiteralPath (Join-Path $RepoRoot $_)) })
    if ($missingSrc.Count) { Fail ("required source files are missing: " + ($missingSrc -join ', ')) }
    Note "required sources present ($($requiredSources.Count))"

    Step 'Check versions'
    $version = Resolve-PackageVersion $RepoRoot $ExpectedVersion
    Note "version $version (VBA PPTK_VERSION = PPTKKeys.csproj <Version>)"

    Step 'Source revision'
    $commit = $null; $branch = $null; $dirty = $null; $isGit = $false
    try { $isGit = ((& git -C $RepoRoot rev-parse --is-inside-work-tree 2>$null) -eq 'true') } catch { $isGit = $false }
    if ($isGit) {
        $commit = (& git -C $RepoRoot rev-parse HEAD).Trim()
        $branch = (& git -C $RepoRoot rev-parse --abbrev-ref HEAD).Trim()
        $dirty = [bool](& git -C $RepoRoot status --porcelain --untracked-files=no -- src)
        Note "commit $commit on $branch; uncommitted changes under src: $dirty"
    } elseif ($AllowNoGit) { Note 'not a Git working tree (allowed for tests): commit and branch recorded as null' }
    else { Fail "$RepoRoot is not a Git working tree; the package must record its source commit" }

    Step 'Check the PPTM against the VBA source'
    $pptmPath = Join-Path $RepoRoot 'src\PPTK-WINDOWS-FINAL.pptm'
    $pptmModules = Get-VbaModules (Read-ZipParts $pptmPath) 'src\PPTK-WINDOWS-FINAL.pptm'
    $noSource = @(); $stale = @()
    foreach ($name in ($pptmModules.Keys | Sort-Object)) {
        $file = Find-SourceFile $RepoRoot $name
        if (-not $file) { $noSource += $name; continue }
        $pptmCode = Get-CodeLines ($pptmModules[$name] -split "`r?`n")
        $srcCode = Get-CodeLines ([IO.File]::ReadAllLines($file, [Text.Encoding]::GetEncoding(1252)))
        if ((Get-Comparable $pptmCode) -ne (Get-Comparable $srcCode)) { $stale += "$name <> $($file.Substring($RepoRoot.Length + 1))" }
    }
    if ($noSource.Count) { Fail ("VBA modules in the PPTM without a source file: " + ($noSource -join ', ')) }
    if ($stale.Count) { Fail ("the PPTM is stale; its VBA differs from the source for:`n    " + ($stale -join "`n    ")) }
    $pptmVersion = Get-ModuleVersion $pptmModules
    if ($pptmVersion -ne $version) { Fail "the PPTM contains PPTK_VERSION '$pptmVersion', expected $version" }
    Note "VBA in the PPTM matches the source ($($pptmModules.Count) modules); PPTK_VERSION $pptmVersion"

    Step 'Check PowerPoint'
    if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Fail 'PowerPoint is running. Close it first: the build never uses a PowerPoint session that is in use.' }
    Note 'PowerPoint is not running'

    $work = Join-Path ([IO.Path]::GetTempPath()) ('PPTK-package-build-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Path $work | Out-Null
    try {
        # ------------------------------------------------ PPAM
        Step 'Build PPTK.ppam'
        $pptmCopy = Join-Path $work 'PPTK-source.pptm'
        Copy-Item -LiteralPath $pptmPath -Destination $pptmCopy
        # Saved under another name: PowerPoint refuses a second file named PPTK.ppam while the
        # registered PPTK.ppam add-in is loaded. It is renamed in the package.
        $builtPpam = Join-Path $work 'PPTK-build.ppam'
        $app = $null; $pres = $null; $startedPid = $null
        try {
            $app = New-Object -ComObject PowerPoint.Application
            $startedPid = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
            $pres = $app.Presentations.Open($pptmCopy, 0, 0, -1)
            # The open document keeps its name after SaveAs to an add-in; it holds the same VBA.
            $runName = "'" + $pres.Name + "'!ModulePPTKBrand.PPTK_GetVersion"
            try { $runVersion = [string][System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, @($runName)) }
            catch { Fail 'PPTK_GetVersion() could not be run' }
            if ($runVersion -ne $version) { Fail "PPTK_GetVersion() returns '$runVersion', expected $version" }
            $pres.SaveAs($builtPpam, 30)   # ppSaveAsOpenXMLAddin
            $pres.Saved = -1; $pres.Close()
            Note "saved with PowerPoint; PPTK_GetVersion() returned $runVersion"
        } finally {
            if ($pres) { try { $pres.Saved = -1; $pres.Close() } catch { }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
            if ($app) { try { $app.Quit() } catch { }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
            $pres = $null; $app = $null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
            for ($i = 0; $i -lt 40 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 500 }
            # Only the instance this build started, and only if it did not exit by itself.
            foreach ($id in @($startedPid)) { if ($id -and (Get-Process -Id $id -ErrorAction SilentlyContinue)) { Stop-Process -Id $id -Force; Note "closed the PowerPoint instance this build started (pid $id)" } }
        }
        if (-not (Test-Path -LiteralPath $builtPpam)) { Fail 'PowerPoint did not write the add-in' }
        $ppamParts = Read-ZipParts $builtPpam
        $images = Test-PpamStructure $ppamParts 'the built add-in' $RepoRoot
        $ppamModules = Get-VbaModules $ppamParts 'the built add-in'
        $ppamVersion = Get-ModuleVersion $ppamModules
        if ($ppamVersion -ne $version) { Fail "the built add-in contains PPTK_VERSION '$ppamVersion', expected $version" }
        $differs = @($pptmModules.Keys | Where-Object { -not $ppamModules.ContainsKey($_) -or (Get-Comparable (Get-CodeLines ($ppamModules[$_] -split "`r?`n"))) -ne (Get-Comparable (Get-CodeLines ($pptmModules[$_] -split "`r?`n"))) })
        if ($differs.Count -or $ppamModules.Count -ne $pptmModules.Count) { Fail ("the built add-in's VBA differs from the PPTM: " + ($differs -join ', ')) }
        $fingerprint = Get-ContentFingerprint $ppamParts $ppamModules
        Note "add-in OK: $($ppamParts.Count) parts, $images ribbon images (= source), $($ppamModules.Count) VBA modules (= PPTM), PPTK_VERSION $ppamVersion"

        # ------------------------------------------------ PPTKKeys
        Step 'Build PPTKKeys (dotnet publish, UseAppHost=false)'
        if (-not $DotnetPath) {
            foreach ($cand in @((Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet\dotnet.exe'), 'C:\Program Files\dotnet\dotnet.exe')) { if (Test-Path -LiteralPath $cand) { $DotnetPath = $cand; break } }
            if (-not $DotnetPath) { $cmd = Get-Command dotnet.exe -ErrorAction SilentlyContinue; if ($cmd) { $DotnetPath = $cmd.Source } }
        }
        if (-not $DotnetPath -or -not (Test-Path -LiteralPath $DotnetPath)) { Fail 'dotnet.exe (.NET 10 SDK) not found' }
        $keysOut = Join-Path $work 'PPTKKeys'
        $publish = & $DotnetPath publish (Join-Path $RepoRoot 'src\PPTKKeys\windows\PPTKKeys.csproj') -c Release -o $keysOut -p:UseAppHost=false --nologo -v q 2>&1
        if ($LASTEXITCODE -ne 0) { Fail ("dotnet publish failed:`n" + ($publish -join "`n")) }
        $exes = @(Get-ChildItem -LiteralPath $keysOut -Recurse -Filter *.exe)
        if ($exes.Count) { Fail ("the PPTKKeys build produced an EXE, which is never packaged or run: " + (($exes | ForEach-Object { $_.Name }) -join ', ')) }
        $missingKeys = @($KeysFiles | Where-Object { -not (Test-Path -LiteralPath (Join-Path $keysOut $_)) })
        if ($missingKeys.Count) { Fail ("the PPTKKeys build is missing: " + ($missingKeys -join ', ')) }
        $dllPath = Join-Path $keysOut 'PPTKKeys.dll'
        $asm = [Reflection.AssemblyName]::GetAssemblyName($dllPath).Version.ToString(3)
        $cliVersion = ([string](& $DotnetPath exec $dllPath --version 2>&1 | Select-Object -First 1)).Trim()
        if ($asm -ne $version -or $cliVersion -ne $version) { Fail "PPTKKeys version mismatch: assembly $asm, --version '$cliVersion', expected $version" }
        $releaseDll = Join-Path $RepoRoot 'src\PPTKKeys\windows\release\win-x64\PPTKKeys.dll'
        if ((Test-Path -LiteralPath $releaseDll) -and (Get-Sha256 $releaseDll) -eq (Get-Sha256 $dllPath)) { Fail 'the built PPTKKeys.dll is identical to the stale release\win-x64 binary' }
        $framework = (Get-Content -LiteralPath (Join-Path $keysOut 'PPTKKeys.runtimeconfig.json') -Raw | ConvertFrom-Json).runtimeOptions.framework
        Note "PPTKKeys.dll assembly $asm, --version $cliVersion; runtime $($framework.name) $($framework.version)"

        # ------------------------------------------------ assemble
        Step 'Assemble package'
        New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
        $pkg = Join-Path (Resolve-Path -LiteralPath $OutputDir).Path $PackageFolderName
        if (Test-Path -LiteralPath $pkg) {
            $isOurs = $false
            $old = Join-Path $pkg 'version.json'
            if (Test-Path -LiteralPath $old) { try { $isOurs = ((Get-Content -LiteralPath $old -Raw | ConvertFrom-Json).packageName -eq $PackageName) } catch { } }
            if (-not $isOurs) { Fail "$pkg exists and is not a PPTK package; not deleting it" }
            Remove-Item -LiteralPath $pkg -Recurse -Force
        }
        New-Item -ItemType Directory -Path (Join-Path $pkg 'payload\PPTKKeys') | Out-Null
        Copy-Item -LiteralPath $builtPpam -Destination (Join-Path $pkg 'payload\PPTK.ppam')
        foreach ($f in $KeysFiles) { Copy-Item -LiteralPath (Join-Path $keysOut $f) -Destination (Join-Path $pkg "payload\PPTKKeys\$f") }
        foreach ($s in $PackageScripts) { Copy-Item -LiteralPath (Join-Path $RepoRoot "src\Deploy\$s") -Destination (Join-Path $pkg $s) }

        $files = New-Object System.Collections.ArrayList
        $add = {
            param($rel, $component, $deterministic, $extra)
            $p = Join-Path $pkg ($rel.Replace('/', '\'))
            $entry = [ordered]@{ path = $rel; component = $component; size = (Get-Item -LiteralPath $p).Length; sha256 = (Get-Sha256 $p); deterministic = $deterministic }
            if ($extra) { foreach ($k in $extra.Keys) { $entry[$k] = $extra[$k] } }
            [void]$files.Add($entry)
        }
        & $add 'payload/PPTK.ppam' 'addin' $false ([ordered]@{ contentFingerprint = $fingerprint; vbaVersion = $ppamVersion })
        & $add 'payload/PPTKKeys/PPTKKeys.dll' 'pptkkeys' $true ([ordered]@{ assemblyVersion = $asm })
        & $add 'payload/PPTKKeys/PPTKKeys.deps.json' 'pptkkeys' $true $null
        & $add 'payload/PPTKKeys/PPTKKeys.runtimeconfig.json' 'pptkkeys' $true $null
        foreach ($s in $PackageScripts) { & $add $s 'script' $true $null }
        $runtimeInfo = [ordered]@{}
        foreach ($rid in $Runtimes.Keys) {
            Step "Bundle the private .NET runtime $rid"
            $rt = Test-Runtime $Runtimes[$rid] $rid $framework.name $framework.version
            $dst = Join-Path $pkg "payload\runtime\$rid"
            New-Item -ItemType Directory -Force -Path $dst | Out-Null
            Copy-Item -Path (Join-Path $Runtimes[$rid] '*') -Destination $dst -Recurse
            foreach ($f in @(Get-ChildItem -LiteralPath $dst -Recurse -File | Sort-Object FullName)) { & $add ("payload/runtime/$rid/" + $f.FullName.Substring($dst.Length + 1).Replace('\', '/')) "runtime-$rid" $true $null }
            $runtimeInfo[$rid] = [ordered]@{ version = $rt.Version; fingerprint = (Get-RuntimeFingerprint $dst); source = "Microsoft dotnet-runtime-$($rt.Version)-$rid.zip" }
            Note "$rid`: .NET $($rt.Version), $($rt.Signed) binaries signed by Microsoft"
        }

        $manifest = [ordered]@{
            schemaVersion  = 1
            packageName    = $PackageName
            packageVersion = $version
            build          = [ordered]@{
                timestampUtc    = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
                sourceCommit    = $commit
                sourceBranch    = $branch
                sourceDirty     = $dirty
                reproducibility = 'Files with deterministic=true are byte-identical for the same source. PPTK.ppam bytes change on every PowerPoint save (ppt/vbaProject.bin); compare its contentFingerprint: SHA-256 over every other add-in part and the VBA source of every module.'
            }
            requirements   = [ordered]@{
                os            = 'Windows'
                office        = 'Microsoft PowerPoint for Windows (desktop, Microsoft 365). Tested on 64-bit; the add-in has no Win32 Declare statements.'
                dotnetRuntime = [ordered]@{ name = $framework.name; minimumVersion = $framework.version; rollForward = 'latest patch within the same major version (.NET default)' }
                pptkkeysHost  = 'dotnet.exe exec PPTKKeys.dll (framework-dependent, AnyCPU; the x64 dotnet host on x64 Windows). No PPTKKeys.exe.'
            }
            files          = $files
        }
        if ($runtimeInfo.Count) {
            $manifest['runtimes'] = $runtimeInfo
            $manifest.requirements['dotnetRuntime']['bundled'] = 'Private official Microsoft .NET runtimes for win-x64 and win-arm64 are included; install.ps1 uses the one for the device architecture. No installed .NET is needed.'
        }
        [IO.File]::WriteAllText((Join-Path $pkg 'version.json'), ($manifest | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
    } finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }

    $m = Test-Package $pkg

    # Intune runs a custom detection script on its own, without parameters and without the package
    # content, so this copy carries the package's version.json (Base64). It sits next to the
    # package, not inside it: version.json cannot list the hash of a file that contains it.
    Step 'Generate the Intune detection script'
    $template = [IO.File]::ReadAllText((Join-Path $RepoRoot 'src\Deploy\detect.ps1'))
    $marker = "`$EmbeddedManifestBase64 = ''"
    if (([regex]::Matches($template, [regex]::Escape($marker))).Count -ne 1) { Fail "src\Deploy\detect.ps1 must contain the line $marker exactly once" }
    # A universal manifest lists ~400 runtime files; the detection script needs only their fingerprints.
    $fullManifest = Get-Content -LiteralPath (Join-Path $pkg 'version.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($fullManifest.PSObject.Properties.Name -contains 'runtimes') {
        $fullManifest.files = @($fullManifest.files | Where-Object { $_.path -notmatch '^payload/runtime/' })
        $manifestBytes = (New-Object Text.UTF8Encoding($false)).GetBytes(($fullManifest | ConvertTo-Json -Depth 6))
    } else { $manifestBytes = [IO.File]::ReadAllBytes((Join-Path $pkg 'version.json')) }
    $b64 = [Convert]::ToBase64String($manifestBytes)
    $detectOut = Join-Path (Split-Path -Parent $pkg) 'PPTK-detect.ps1'
    [IO.File]::WriteAllText($detectOut, $template.Replace($marker, "`$EmbeddedManifestBase64 = '$b64'"), (New-Object Text.UTF8Encoding($false)))
    $embedded = [regex]::Match([IO.File]::ReadAllText($detectOut), "\`$EmbeddedManifestBase64 = '([A-Za-z0-9+/=]+)'")
    if (-not $embedded.Success -or (Get-Sha256Bytes ([Convert]::FromBase64String($embedded.Groups[1].Value))) -ne (Get-Sha256Bytes $manifestBytes)) { Fail 'the generated detection script does not carry this package''s version.json' }
    Note "PPTK-detect.ps1 (embedded version.json, SHA-256 $(Get-Sha256 $detectOut)) -> $detectOut"

    Write-Host "PACKAGE OK: $($m.packageName) $($m.packageVersion) -> $pkg"
    exit 0
} catch {
    Write-Host ($_.Exception.Message) -ForegroundColor Red
    exit 1
}
