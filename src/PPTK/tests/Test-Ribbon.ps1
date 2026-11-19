<#
.SYNOPSIS
  Automated PowerPoint test for the PPTK ribbon (customUI on the Office 2010+ 2009/07 schema, with context menus).

.DESCRIPTION
  Package checks: the root relationship is the 2007 ui/extensibility type, the ribbon XML is on the 2009/07
  schema, and the right-click Select menu is on the five context menus.
  PowerPoint checks: a copy of the PPTM is made in %TEMP% with only the tab label changed to "PPTK TEST" (so it
  cannot be mistaken for the tab of an installed PPTK), opened with a window in a private PowerPoint. Then:
  the ribbon onLoad ran (ModulePPTKAlign.PPTK_RibbonLoaded), the groups and controls are on screen (UI
  Automation, window maximized so no group collapses), and clicking "Left" runs its macro (aligns to the first
  selected shape). The right-click menu itself is checked by hand (UI Automation cannot open it without input).
  The copy is deleted afterwards. Exit code 0 when every case passes.
#>
param([string]$Pptm = '')

$ErrorActionPreference = 'Stop'
if (-not $Pptm) { $Pptm = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path 'PPTK-WINDOWS-FINAL.pptm' }
$Pptm = (Resolve-Path -LiteralPath $Pptm).Path
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Host 'SKIPPED: close PowerPoint first'; exit 1 }
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem, UIAutomationClient, UIAutomationTypes
$results = New-Object System.Collections.ArrayList
function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) { [void]$results.Add($ok); Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail) }
function ReadEntry($zip, [string]$name) { $e = $zip.GetEntry($name); if (-not $e) { return $null }; $r = New-Object IO.StreamReader($e.Open()); try { $r.ReadToEnd() } finally { $r.Dispose() } }

# ---- package
$z = [IO.Compression.ZipFile]::OpenRead($Pptm)
try { $rels = ReadEntry $z '_rels/.rels'; $ui = ReadEntry $z 'customUI/customUI.xml' } finally { $z.Dispose() }
Record 'K1' 'root relationship is the 2007 ui/extensibility type' ($rels -match 'Type="http://schemas\.microsoft\.com/office/2007/relationships/ui/extensibility"') ''
$doc = [xml]$ui
Record 'K2' 'ribbon XML is on the 2009/07 schema' ($doc.DocumentElement.NamespaceURI -eq 'http://schemas.microsoft.com/office/2009/07/customui') $doc.DocumentElement.NamespaceURI
$ns = New-Object Xml.XmlNamespaceManager($doc.NameTable); $ns.AddNamespace('c', 'http://schemas.microsoft.com/office/2009/07/customui')
$menus = @($doc.SelectNodes('//c:contextMenu', $ns) | ForEach-Object { $_.idMso })
$wanted = 'ContextMenuFrame', 'ContextMenuShape', 'ContextMenuObjectsGroup', 'ContextMenuPicture', 'ContextMenuGraphic'
$okMenus = $true
foreach ($m in $wanted) {
    $sel = $doc.SelectSingleNode("//c:contextMenu[@idMso='$m']/c:menu[@label='Select']", $ns)
    if (-not $sel -or @($sel.SelectNodes('c:dynamicMenu', $ns)).Count -ne 8 -or @($sel.SelectNodes('c:button', $ns)).Count -ne 2) { $okMenus = $false }
}
Record 'K3' 'Select menu (2 buttons, 8 Select By submenus) on the five context menus' ($okMenus -and $menus.Count -eq 5) ($menus -join ', ')
$ctxControls = @($doc.SelectNodes('//c:contextMenu//*[@id]', $ns))
$ctxImages = @($ctxControls | ForEach-Object { $_.GetAttribute('image') } | Sort-Object -Unique)
$z = [IO.Compression.ZipFile]::OpenRead($Pptm)
try { $uiRels = ReadEntry $z 'customUI/_rels/customUI.xml.rels'; $imgParts = @($z.Entries | Where-Object { $_.FullName -like 'customUI/images/*' } | ForEach-Object { $_.Name }) } finally { $z.Dispose() }
$broken = @($ctxImages | Where-Object { -not $_ -or $uiRels -notmatch "Id=""$_""[^>]*Target=""images/$_\.png""" -or $imgParts -notcontains "$_.png" })
Record 'K5' 'Select menu: every item has a PPTK icon (no Office icon), each with a relationship and a PNG in the file' ($ctxControls.Count -eq 55 -and @($ctxControls | Where-Object { -not $_.GetAttribute('image') -or $_.GetAttribute('imageMso') }).Count -eq 0 -and $ctxImages.Count -eq 11 -and $broken.Count -eq 0) ("icons: " + ($ctxImages -join ', ') + $(if ($broken) { ' | broken: ' + ($broken -join ', ') } else { '' }))
$ids = @($doc.SelectNodes('//*[@id]') | ForEach-Object { $_.id })
Record 'K4' 'control ids are unique' ($ids.Count -eq @($ids | Sort-Object -Unique).Count) "$($ids.Count) ids"

# ---- PowerPoint: a copy with only the tab label changed
$copy = Join-Path $env:TEMP ('PPTK-ribbon-test-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.pptm')
Copy-Item -LiteralPath $Pptm $copy
$z = [IO.Compression.ZipFile]::Open($copy, 'Update')
try {
    $text = ReadEntry $z 'customUI/customUI.xml'
    $text = $text.Replace('<tab id="InstrumentaPowerpointToolbar" label="PPTK">', '<tab id="InstrumentaPowerpointToolbar" label="PPTK TEST">')
    $z.GetEntry('customUI/customUI.xml').Delete(); $n = $z.CreateEntry('customUI/customUI.xml')
    $w = New-Object IO.StreamWriter($n.Open(), (New-Object Text.UTF8Encoding($false))); $w.Write($text); $w.Dispose()
} finally { $z.Dispose() }

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$pres = $null
try {
    $pres = $app.Presentations.Open($copy, 0, 0, -1)
    $app.WindowState = 3
    $s = $pres.Slides.Add($pres.Slides.Count + 1, 12)
    $app.ActiveWindow.ViewType = 9
    $app.ActiveWindow.View.GotoSlide($s.SlideIndex)
    $a = $s.Shapes.AddShape(1, 100, 100, 80, 60); $a.Name = 'A'
    $b = $s.Shapes.AddShape(1, 300, 220, 120, 40); $b.Name = 'B'
    Start-Sleep 4
    $loaded = [System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@("'" + $pres.Name + "'!ModulePPTKAlign.PPTK_RibbonLoaded"))
    Record 'L1' 'ribbon onLoad ran' ([bool]$loaded) ''

    $pp = (Get-Process POWERPNT | Sort-Object StartTime | Select-Object -Last 1).Id
    $root = [Windows.Automation.AutomationElement]::RootElement
    $pc = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $pp)
    function FindName($name, $type = $null) { @($root.FindAll([Windows.Automation.TreeScope]::Descendants, $pc) | Where-Object { $_.Current.Name -eq $name -and (-not $type -or $_.Current.LocalizedControlType -eq $type) }) | Select-Object -First 1 }
    $tab = FindName 'PPTK TEST' 'tab item'
    Record 'U1' 'tab on the ribbon' ($null -ne $tab) ''
    if ($tab) { ($tab.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern)).Select(); Start-Sleep 2 }
    $names = @($root.FindAll([Windows.Automation.TreeScope]::Descendants, $pc) | ForEach-Object { $_.Current.Name })
    $expected = 'Design', 'Reuse', 'Insert', 'Text', 'Align', 'Space', 'Rows / TK', 'Tools', 'Shape', 'Left', 'Center', 'Right', 'Top', 'Middle', 'Bottom', 'Central', 'Process', 'Matrix', 'Save Slides', 'Email Slides', 'Fix Slide Colors', 'To selection'
    $missing = @($expected | Where-Object { $names -notcontains $_ })
    Record 'U2' 'groups and controls present' ($missing.Count -eq 0) ('missing: ' + ($missing -join ', '))
    $b.Select(-1); $a.Select(0); $bLeft = $b.Left
    $left = FindName 'Left' 'button'
    if ($left) { ($left.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern)).Invoke(); Start-Sleep 2 }
    Record 'U3' 'ribbon "Left" runs its macro (A to first selected B)' ($left -and [Math]::Abs($a.Left - $bLeft) -lt 0.05) "A.Left=$($a.Left) B.Left=$bLeft"
} catch {
    Record 'X' 'unexpected error' $false $_.Exception.Message
} finally {
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    for ($i = 0; $i -lt 10 -and (Test-Path $copy); $i++) { try { Remove-Item $copy -Force } catch { Start-Sleep 1 } }
}
$failed = @($results | Where-Object { -not $_ }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
