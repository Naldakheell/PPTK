<#
.SYNOPSIS
  Checks the Replace Fonts entry points: the PPTK ribbon Text group button and Smart Bar command 62.

.DESCRIPTION
  Phase A put PowerPoint's native Replace Fonts in the PPTK ribbon. Phase B (2026-11-18) replaced it there with
  the PPTK Replace Fonts window (Arabic and theme fonts; engine tested by Test-ReplaceFontsB.ps1). PowerPoint's
  own command on the Home tab is untouched.
  Checks: one PPTK Replace Fonts button (tag CMD REPLACE FONTS) in the Text group next to the Font box, no native
  copy in the PPTK ribbon; PowerPoint's own command still enabled; the Smart Bar has command 62. The window is
  modal and is not opened here. Same private-PowerPoint rules as Test-Align.ps1. Exit code 0 when every case passes.
#>
param([string]$Pptm = '')

$ErrorActionPreference = 'Stop'
if (-not $Pptm) { $Pptm = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path 'PPTK-WINDOWS-FINAL.pptm' }
$Pptm = (Resolve-Path -LiteralPath $Pptm).Path
$results = New-Object System.Collections.ArrayList
function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) { [void]$results.Add($ok); Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail) }

# 1. ribbon XML inside the PPTM
Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [IO.Compression.ZipFile]::OpenRead($Pptm)
try { $r = New-Object IO.StreamReader($z.GetEntry('customUI/customUI.xml').Open()); [xml]$ui = $r.ReadToEnd(); $r.Dispose() } finally { $z.Dispose() }
$ns = New-Object Xml.XmlNamespaceManager($ui.NameTable); $ns.AddNamespace('c', $ui.DocumentElement.NamespaceURI)
$native = @($ui.SelectNodes("//c:*[@idMso='FontsReplaceFonts']", $ns))
$all = @($ui.SelectNodes("//c:*[@tag='CMD REPLACE FONTS']", $ns))
$inBox = @($ui.SelectNodes("//c:group[@id='Text']/c:box[@id='FontBox1']/c:button[@id='PPTKReplaceFonts'][@tag='CMD REPLACE FONTS'][@onAction='PPTK_RibbonShortcut']", $ns))
Record 'X1' 'one Replace Fonts button, next to the Font box in the Text group, opening the PPTK window; no native copy in the PPTK ribbon' ($all.Count -eq 1 -and $inBox.Count -eq 1 -and $native.Count -eq 0) "PPTK buttons $($all.Count), in FontBox1 $($inBox.Count), native idMso $($native.Count)"

# 2. PowerPoint enables the command; Smart Bar count
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Host 'SKIPPED (PowerPoint part): close PowerPoint first'; exit 1 }
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$addin = $null; $pres = $null
try {
    $addin = $app.Presentations.Open($Pptm, -1, 0, 0)
    $pres = $app.Presentations.Add(-1); [void]$pres.Slides.Add(1, 12)
    Record 'P1' "PowerPoint's own Replace Fonts (Home tab) is still available" ([bool]$app.CommandBars.GetEnabledMso('FontsReplaceFonts')) ''
    $count = [System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@("'" + $addin.Name + "'!ModulePPTKSmartBar.PPTK_CommandCount"))
    Record 'S1' 'Smart Bar lists at least 62 commands (Replace Fonts is 62)' ($count -ge 62) "count $count"
} catch {
    Record 'X' 'unexpected error' $false $_.Exception.Message
} finally {
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    try { if ($addin) { $addin.Close() } } catch { }
    try { $app.Visible = 0 } catch { }
    $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
$failed = @($results | Where-Object { -not $_ }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
