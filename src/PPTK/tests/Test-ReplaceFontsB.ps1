<#
.SYNOPSIS
  Automated PowerPoint test for PPTK Replace Fonts Phase B (ModulePPTKReplaceFonts.PPTK_ReplaceFontsRun).

.DESCRIPTION
  A private PowerPoint; the PPTM given by -Pptm opened without a window; test decks built here and saved in
  %TEMP% (deleted afterwards). Each case builds a fresh deck, calls PPTK_ReplaceFontsRun "source|target|scope"
  through Application.Run (InvokeMember), and checks text runs, theme fonts (ThemeFontScheme.Save), the saved
  XML (theme-linked text must stay linked), new slides, save/close/reopen, unrelated fonts and formatting, one
  Undo, and the backup copy (Documents\PPTK Font Backups; each backup made by this test is deleted).
  Exit code 0 when every case passes.
#>
param([string]$Pptm = '')

$ErrorActionPreference = 'Stop'
if (-not $Pptm) { $Pptm = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path 'PPTK-WINDOWS-FINAL.pptm' }
$Pptm = (Resolve-Path -LiteralPath $Pptm).Path
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Host 'SKIPPED: close PowerPoint first'; exit 1 }
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$results = New-Object System.Collections.ArrayList
function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) { [void]$results.Add($ok); Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail) }
$work = Join-Path $env:TEMP ('pptk-fonts-test-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory $work | Out-Null
$backupDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PPTK Font Backups'
$backupDirExisted = Test-Path $backupDir
$ar = [string]::new([char[]](0x0645, 0x0631, 0x062D, 0x0628, 0x0627))   # 5 Arabic letters
$T = 'Segoe UI'

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$addin = $null
try {
    $addin = $app.Presentations.Open($Pptm, -1, 0, 0)
    $prefix = "'" + $addin.Name + "'!"
    function Replace([string]$spec) { [string][System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@(($prefix + 'ModulePPTKReplaceFonts.PPTK_ReplaceFontsRun'), $spec)) }
    function Field([string]$result, [string]$name) { foreach ($p in $result -split ';') { if ($p.StartsWith($name + '=')) { return $p.Substring($name.Length + 1) } }; '' }
    function DropBackup([string]$result) { $b = Field $result 'backup'; if ($b -and (Test-Path -LiteralPath $b)) { Remove-Item -LiteralPath $b -Force; return $true }; $false }

    # ---- the test deck
    function NewDeck([string]$name) {
        $p = $app.Presentations.Add(-1)
        $app.ActiveWindow.ViewType = 9
        $s = $p.Slides.Add(1, 12)
        function Box($n, $text, $x, $y) { $b = $s.Shapes.AddTextbox(1, $x, $y, 260, 30); $b.Name = $n; $b.TextFrame2.TextRange.Text = $text; $b }
        $null = Box 'AR-theme' $ar 20 10
        $b = Box 'AR-explicit' $ar 20 50; $b.TextFrame2.TextRange.Font.NameComplexScript = 'Arial'
        $b = Box 'AR-tahoma' $ar 20 90; $b.TextFrame2.TextRange.Font.NameComplexScript = 'Tahoma'
        $null = Box 'EN-theme' 'Hello' 20 130
        $b = Box 'EN-arial' 'Hello' 20 170; $b.TextFrame2.TextRange.Font.Name = 'Arial'
        $b = Box 'EN-calibri' 'Hello' 20 210; $b.TextFrame2.TextRange.Font.Name = 'Calibri'
        $b = Box 'MIX' ($ar + ' Hello') 300 10; $r = $b.TextFrame2.TextRange
        $r.Characters(1, 5).Font.NameComplexScript = 'Arial'; $r.Characters(1, 5).Font.Bold = -1; $r.Characters(1, 5).Font.Size = 22; $r.Characters(1, 5).Font.Fill.ForeColor.RGB = 255
        $r.Characters(7, 5).Font.Name = 'Calibri'; $r.Characters(7, 5).Font.Italic = -1
        $r.ParagraphFormat.Alignment = 3
        $b = Box 'BOTH' ($ar + ' Hello') 300 50; $b.TextFrame2.TextRange.Font.Name = 'Arial'; $b.TextFrame2.TextRange.Font.NameComplexScript = 'Arial'
        $b = Box 'MULTI' 'One Two Three' 300 90; $r = $b.TextFrame2.TextRange
        $r.Characters(1, 3).Font.Name = 'Arial'; $r.Characters(1, 3).Font.Size = 14; $r.Characters(5, 3).Font.Name = 'Calibri'; $r.Characters(5, 3).Font.Size = 16; $r.Characters(9, 5).Font.Name = 'Arial'; $r.Characters(9, 5).Font.Size = 18
        $t = $s.Shapes.AddTable(1, 2, 300, 140, 300, 40); $t.Name = 'TABLE'
        $c = $t.Table.Cell(1, 1).Shape.TextFrame2.TextRange; $c.Text = $ar; $c.Font.NameComplexScript = 'Arial'
        $c = $t.Table.Cell(1, 2).Shape.TextFrame2.TextRange; $c.Text = 'Cell'; $c.Font.Name = 'Arial'
        $g1 = Box 'G1' $ar 300 200; $g1.TextFrame2.TextRange.Font.NameComplexScript = 'Arial'
        $g2 = Box 'G2' 'Grouped' 300 240; $g2.TextFrame2.TextRange.Font.Name = 'Arial'
        $grp = $s.Shapes.Range([object[]]@('G1', 'G2')).Group(); $grp.Name = 'GROUP'
        $n = $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame2.TextRange; $n.Text = $ar + ' Notes'; $n.Font.Name = 'Arial'; $n.Font.NameComplexScript = 'Arial'
        $m = $p.SlideMaster.Shapes.AddTextbox(1, 600, 480, 200, 30); $m.Name = 'MASTER-BOX'; $m.TextFrame2.TextRange.Text = 'Master'; $m.TextFrame2.TextRange.Font.Name = 'Arial'
        $l = $p.SlideMaster.CustomLayouts.Item(1).Shapes.AddTextbox(1, 600, 440, 200, 30); $l.Name = 'LAYOUT-BOX'; $l.TextFrame2.TextRange.Text = $ar; $l.TextFrame2.TextRange.Font.NameComplexScript = 'Arial'
        $nm = $p.NotesMaster.Shapes.AddTextbox(1, 20, 20, 200, 30); $nm.Name = 'NOTESMASTER-BOX'; $nm.TextFrame2.TextRange.Text = 'NM'; $nm.TextFrame2.TextRange.Font.Name = 'Arial'
        $s2 = $p.Slides.Add(2, 2); $s2.Shapes.Item(1).TextFrame2.TextRange.Text = $ar; $s2.Shapes.Item(2).TextFrame2.TextRange.Text = $ar + ' Body'
        $file = Join-Path $work "$name.pptx"; $p.SaveAs($file, 24)
        $app.ActiveWindow.View.GotoSlide(1)
        $p
    }
    function Sh($p, [string]$n) { $p.Slides.Item(1).Shapes.Item($n) }
    function Cs($p, [string]$n) { (Sh $p $n).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript }
    function Lt($p, [string]$n) { (Sh $p $n).TextFrame2.TextRange.Font.Name }
    function ThemeFonts($p) {
        $f = Join-Path $work ('scheme-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.xml'); $p.SlideMaster.Theme.ThemeFontScheme.Save($f); $x = [IO.File]::ReadAllText($f)
        $mj = [regex]::Match($x, '<a:majorFont>.*?</a:majorFont>', 'Singleline').Value; $mn = [regex]::Match($x, '<a:minorFont>.*?</a:minorFont>', 'Singleline').Value
        [pscustomobject]@{ MajorLatin = [regex]::Match($mj, '<a:latin typeface="([^"]*)"').Groups[1].Value; MajorArab = [regex]::Match($mj, '<a:font script="Arab" typeface="([^"]*)"').Groups[1].Value
                           MinorLatin = [regex]::Match($mn, '<a:latin typeface="([^"]*)"').Groups[1].Value; MinorArab = [regex]::Match($mn, '<a:font script="Arab" typeface="([^"]*)"').Groups[1].Value }
    }
    # Everything except fonts: text, geometry, size, bold, italic, color, alignment of every run on slide 1.
    function Look($p) {
        $out = @()
        foreach ($x in @($p.Slides.Item(1).Shapes)) {
            $out += "$($x.Name)@$([Math]::Round($x.Left,2)),$([Math]::Round($x.Top,2)),$([Math]::Round($x.Width,2)),$([Math]::Round($x.Height,2))"
            if ($x.HasTextFrame -and $x.Type -ne 6) { $r = $x.TextFrame2.TextRange; $out += "'$($r.Text)' al$($r.ParagraphFormat.Alignment)"; for ($i = 1; $i -le $r.Runs().Count; $i++) { $u = $r.Runs($i); $out += "s$($u.Font.Size) b$($u.Font.Bold) i$($u.Font.Italic) c$($u.Font.Fill.ForeColor.RGB)" } }
        }
        $out -join ' | '
    }
    function SavedShapeXml([string]$file, [string]$entry, [string]$shapeName) {
        $z = [IO.Compression.ZipFile]::OpenRead($file); try { $r = New-Object IO.StreamReader($z.GetEntry($entry).Open()); $x = $r.ReadToEnd(); $r.Dispose() } finally { $z.Dispose() }
        [regex]::Match($x, "<p:sp>(?:(?!</p:sp>).)*name=""$shapeName"".*?</p:sp>", 'Singleline').Value
    }

    # ================================================================ All text: Arial -> Segoe UI
    $p = NewDeck 'all'; $before = Look $p
    $res = Replace "Arial|$T|All"
    Record 'R0' 'run result: backup made, nothing left, nothing failed' ((Field $res 'left') -eq '0' -and (Field $res 'failed') -eq '0' -and (DropBackup $res)) $res
    Record 'A1' 'Arabic-only text (explicit Arial) replaced' ((Cs $p 'AR-explicit') -eq $T) (Cs $p 'AR-explicit')
    Record 'A2' 'English-only text (explicit Arial) replaced' ((Lt $p 'EN-arial') -eq $T) (Lt $p 'EN-arial')
    $mix = (Sh $p 'MIX').TextFrame2.TextRange
    Record 'A3' 'mixed text: Arabic Arial replaced, English Calibri kept' ($mix.Characters(1, 5).Font.NameComplexScript -eq $T -and $mix.Characters(7, 5).Font.Name -eq 'Calibri') "$($mix.Characters(1,5).Font.NameComplexScript) / $($mix.Characters(7,5).Font.Name)"
    $multi = (Sh $p 'MULTI').TextFrame2.TextRange
    Record 'A4' 'several fonts in one box: only the Arial runs change' ($multi.Characters(1, 3).Font.Name -eq $T -and $multi.Characters(5, 3).Font.Name -eq 'Calibri' -and $multi.Characters(9, 5).Font.Name -eq $T) "$($multi.Characters(1,3).Font.Name) / $($multi.Characters(5,3).Font.Name) / $($multi.Characters(9,5).Font.Name)"
    $tbl = (Sh $p 'TABLE').Table; $grp = Sh $p 'GROUP'
    Record 'A5' 'table cells and grouped shapes' ($tbl.Cell(1, 1).Shape.TextFrame2.TextRange.Font.NameComplexScript -eq $T -and $tbl.Cell(1, 2).Shape.TextFrame2.TextRange.Font.Name -eq $T -and $grp.GroupItems.Item('G1').TextFrame2.TextRange.Font.NameComplexScript -eq $T -and $grp.GroupItems.Item('G2').TextFrame2.TextRange.Font.Name -eq $T) ''
    $notes = $p.Slides.Item(1).NotesPage.Shapes.Placeholders.Item(2).TextFrame2.TextRange
    $okA6 = $notes.Characters(1, 5).Font.NameComplexScript -eq $T -and $notes.Characters(7, 5).Font.Name -eq $T -and $p.SlideMaster.Shapes.Item('MASTER-BOX').TextFrame2.TextRange.Font.Name -eq $T -and $p.SlideMaster.CustomLayouts.Item(1).Shapes.Item('LAYOUT-BOX').TextFrame2.TextRange.Font.NameComplexScript -eq $T -and $p.NotesMaster.Shapes.Item('NOTESMASTER-BOX').TextFrame2.TextRange.Font.Name -eq $T
    Record 'A6' 'notes page, slide master, layout and notes master' $okA6 ''
    $th = ThemeFonts $p
    Record 'A7' 'theme: minor Arabic (Arial) replaced; major Arabic and Latin fonts untouched' ($th.MinorArab -eq $T -and $th.MajorArab -eq 'Times New Roman' -and $th.MinorLatin -eq 'Aptos' -and $th.MajorLatin -eq 'Aptos Display') "$($th.MinorArab) / $($th.MajorArab) / $($th.MinorLatin) / $($th.MajorLatin)"
    Record 'A8' 'theme-linked Arabic text follows the theme' ((Cs $p 'AR-theme') -eq $T -and ((Sh $p 'BOTH').TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript) -eq $T) (Cs $p 'AR-theme')
    $n1 = $p.Slides.Add(3, 2); $n1.Shapes.Item(2).TextFrame2.TextRange.Text = $ar + ' Hello'
    Record 'A9' 'new slide after replacement: Arabic uses the new font' ($n1.Shapes.Item(2).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript -eq $T) $n1.Shapes.Item(2).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript
    $n1.Delete()
    Record 'A10' 'unrelated fonts kept (Tahoma, Calibri, theme Latin), formatting/text/positions identical' ((Cs $p 'AR-tahoma') -eq 'Tahoma' -and (Lt $p 'EN-calibri') -eq 'Calibri' -and (Lt $p 'EN-theme') -eq 'Aptos' -and (Look $p) -eq $before) ''
    $file = $p.FullName; $p.Save(); $p.Close()
    $themeLinked = SavedShapeXml $file 'ppt/slides/slide1.xml' 'AR-theme'
    Record 'A11' 'saved file: theme-linked Arabic text has no explicit font (still linked)' ($themeLinked -and $themeLinked -notmatch '<a:cs ') ''
    $p = $app.Presentations.Open($file, 0, 0, -1); $app.ActiveWindow.View.GotoSlide(1)
    $th = ThemeFonts $p
    $n2 = $p.Slides.Add(3, 2); $n2.Shapes.Item(2).TextFrame2.TextRange.Text = $ar
    Record 'A12' 'after save, close and reopen: theme, text and new slides keep the new font' ($th.MinorArab -eq $T -and (Cs $p 'AR-explicit') -eq $T -and (Lt $p 'EN-arial') -eq $T -and (Cs $p 'AR-theme') -eq $T -and $n2.Shapes.Item(2).TextFrame2.TextRange.Font.NameComplexScript -eq $T) "theme Arab $($th.MinorArab)"
    $p.Saved = -1; $p.Close()

    # ================================================================ Arabic only
    $p = NewDeck 'arabic'
    $res = Replace "Arial|$T|Arabic"; [void](DropBackup $res)
    $both = (Sh $p 'BOTH').TextFrame2.TextRange
    Record 'B1' 'Arabic only: Arabic Arial replaced, English Arial kept' ((Cs $p 'AR-explicit') -eq $T -and (Lt $p 'EN-arial') -eq 'Arial' -and $both.Characters(1, 5).Font.NameComplexScript -eq $T -and $both.Characters(7, 5).Font.Name -eq 'Arial' -and (ThemeFonts $p).MinorArab -eq $T) "$($both.Characters(1,5).Font.NameComplexScript) / $($both.Characters(7,5).Font.Name)"
    $p.Saved = -1; $p.Close()

    # ================================================================ English only
    $p = NewDeck 'english'
    $res = Replace "Arial|$T|English"; [void](DropBackup $res)
    $both = (Sh $p 'BOTH').TextFrame2.TextRange
    Record 'C1' 'English only: English Arial replaced; Arabic Arial and the Arabic theme font kept' ((Lt $p 'EN-arial') -eq $T -and (Cs $p 'AR-explicit') -eq 'Arial' -and $both.Characters(1, 5).Font.NameComplexScript -eq 'Arial' -and $both.Characters(7, 5).Font.Name -eq $T -and (ThemeFonts $p).MinorArab -eq 'Arial' -and (Cs $p 'AR-theme') -eq 'Arial') "$($both.Characters(1,5).Font.NameComplexScript) / $($both.Characters(7,5).Font.Name)"
    $p.Saved = -1; $p.Close()

    # ================================================================ theme fonts: Latin minor and Arabic major
    $p = NewDeck 'latin-theme'
    $res = Replace 'Aptos|Georgia|English'; [void](DropBackup $res)
    $n3 = $p.Slides.Add(3, 2); $n3.Shapes.Item(2).TextFrame2.TextRange.Text = 'New'
    $th = ThemeFonts $p
    Record 'D1' 'Latin theme (minor Aptos -> Georgia): theme, linked text and new slides; major Aptos Display kept' ($th.MinorLatin -eq 'Georgia' -and $th.MajorLatin -eq 'Aptos Display' -and (Lt $p 'EN-theme') -eq 'Georgia' -and $n3.Shapes.Item(2).TextFrame2.TextRange.Font.Name -eq 'Georgia' -and (Lt $p 'EN-calibri') -eq 'Calibri') "$($th.MinorLatin) / $($th.MajorLatin)"
    $n3.Delete(); $file = $p.FullName; $p.Save(); $p.Close()
    Record 'D2' 'saved file: theme-linked English text has no explicit font' ((SavedShapeXml $file 'ppt/slides/slide1.xml' 'EN-theme') -notmatch '<a:latin ') ''
    $p = $app.Presentations.Open($file, 0, 0, -1)
    $res = Replace 'Times New Roman|Sakkal Majalla|Arabic'; [void](DropBackup $res)
    $th = ThemeFonts $p
    Record 'D3' 'Arabic major theme font (titles): Times New Roman -> Sakkal Majalla; minor Arabic kept' ($th.MajorArab -eq 'Sakkal Majalla' -and $th.MinorArab -eq 'Arial' -and $p.Slides.Item(2).Shapes.Item(1).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript -eq 'Sakkal Majalla' -and $p.Slides.Item(2).Shapes.Item(2).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript -eq 'Arial') "$($th.MajorArab) / $($th.MinorArab)"
    $p.Saved = -1; $p.Close()

    # ================================================================ one Undo
    $p = NewDeck 'undo'; $before = Look $p
    $res = Replace "Arial|$T|All"; [void](DropBackup $res)
    $app.CommandBars.ExecuteMso('Undo'); Start-Sleep -Milliseconds 500
    $th = ThemeFonts $p
    Record 'U1' 'one Undo restores text and theme fonts' ((Cs $p 'AR-explicit') -eq 'Arial' -and (Lt $p 'EN-arial') -eq 'Arial' -and (Cs $p 'AR-theme') -eq 'Arial' -and $th.MinorArab -eq 'Arial' -and (Look $p) -eq $before) "AR $(Cs $p 'AR-explicit') EN $(Lt $p 'EN-arial') theme $($th.MinorArab)"
    $p.Saved = -1; $p.Close()

    # ================================================================ guards and the font list
    $p = NewDeck 'guards'
    Record 'E1' 'same source and target: refused, nothing changed' ((Replace 'Arial|arial|All') -like 'error=*' -and (Cs $p 'AR-explicit') -eq 'Arial') ''
    $list = [string][System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@(($prefix + 'ModulePPTKReplaceFonts.PPTK_FontsInUse')))
    $names = $list -split "`n"
    $missing = @(@('Arial', 'Calibri', 'Tahoma', 'Aptos', 'Aptos Display', 'Times New Roman') | Where-Object { $names -notcontains $_ })
    Record 'E2' 'font list has text and theme fonts, Latin and Arabic' ($missing.Count -eq 0) ($names -join ', ')

    $inst = [string][System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@(($prefix + 'ModulePPTKReplaceFonts.PPTK_InstalledFonts')))
    $instNames = @($inst -split "`n" | Where-Object { $_ })
    $bad = @($instNames | Where-Object { $_ -match '\(TrueType\)|\.ttf$|\.otf$|\.ttc$' })
    $sorted = (($instNames | Sort-Object) -join '|') -eq ($instNames -join '|')
    Record 'E6' '"Replace with" lists the installed fonts: many families, real names, sorted, no duplicates' ($instNames.Count -gt 50 -and $instNames -contains 'Arial' -and $instNames -contains 'Tahoma' -and $instNames -contains 'Segoe UI' -and $bad.Count -eq 0 -and $sorted -and @($instNames | Sort-Object -Unique).Count -eq $instNames.Count) "count $($instNames.Count); bad: $($bad -join ', ')"
    # Every family of every font file installed in Windows (machine and user font folders, read independently
    # with .NET) must be in the list, however the font was installed. Fonts Windows shows from elsewhere
    # (fonts inside Microsoft Store app packages) are reported, not required.
    Add-Type -AssemblyName System.Drawing
    $fileFamilies = New-Object System.Collections.Generic.HashSet[string]
    foreach ($dir in "$env:WINDIR\Fonts", "$env:LOCALAPPDATA\Microsoft\Windows\Fonts") {
        foreach ($file in @(Get-ChildItem $dir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.(ttf|otf|ttc)$' })) {
            $pfc = New-Object System.Drawing.Text.PrivateFontCollection
            try { $pfc.AddFontFile($file.FullName); foreach ($fam in $pfc.Families) { [void]$fileFamilies.Add($fam.Name) } } catch { } finally { $pfc.Dispose() }
        }
    }
    $missing = @($fileFamilies | Where-Object { $instNames -notcontains $_ })
    $elsewhere = @((New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name } | Where-Object { $instNames -notcontains $_ -and -not $fileFamilies.Contains($_) })
    # Known limit (option 1, no Word): some weights of variable fonts are named by Windows from rules PPTK does
    # not rebuild; they can be typed. Everything else must be there: at least 97% of the families.
    Record 'E7' 'installed font files in Windows are listed by family name, however installed (>= 97%; missing ones can be typed)' ($fileFamilies.Count -gt 0 -and $missing.Count / $fileFamilies.Count -le 0.03) "families from font files $($fileFamilies.Count), missing $($missing.Count): $(($missing | Select-Object -First 10) -join ', ') | note, fonts from Store app packages (type the name): $(($elsewhere | Select-Object -First 6) -join ', ')"
    # Fonts that Windows lists only by file name (installers often do this) are listed by their real name.
    $byFile = @((Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts').PSObject.Properties | Where-Object { $_.Name -match '\.(ttf|otf)$' })
    $ubuntu = @(Get-ChildItem "$env:WINDIR\Fonts" -Filter 'Ubuntu*.ttf' -ErrorAction SilentlyContinue)
    Record 'E8' "fonts registered only by file name are listed by real name ($($byFile.Count) on this machine; e.g. Ubuntu, Ubuntu Arabic)" ($ubuntu.Count -eq 0 -or ($instNames -contains 'Ubuntu' -and ($instNames | Where-Object { $_ -like 'Ubuntu Arabic*' }).Count -gt 0)) (($instNames | Where-Object { $_ -like 'Ubuntu*' -or $_ -like '29LT*' -or $_ -like 'Helvetica*Arabic*' }) -join ', ')
    $catalog = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\FontCache\*\Catalog\ListAll.Json" -ErrorAction SilentlyContinue | Select-Object -First 1
    Record 'E9' "Office's downloadable fonts are listed too (when Office's font catalog is on the machine)" ($null -eq $catalog -or ($instNames -contains 'Poppins' -and $instNames -contains 'Abadi')) "catalog: $(if ($catalog) { 'yes' } else { 'no' })"
    Record 'E3' 'the two "(Theme Arabic ...)" entries are always offered, at the top of the list' ($names[0] -eq '(Theme Arabic body font)' -and $names[1] -eq '(Theme Arabic heading font)') (($names | Select-Object -First 3) -join ', ')
    # PowerPoint's default theme names its Arabic fonts (heading Times New Roman, body Arial) through script entries.
    $res = Replace '(Theme Arabic heading font)|Sakkal Majalla|Arabic'; [void](DropBackup $res)
    $th = ThemeFonts $p; $s2 = $p.Slides.Item(2)
    Record 'E4' 'default theme, heading entry: Arabic titles change; body Arabic and explicit Arial text untouched' ($th.MajorArab -eq 'Sakkal Majalla' -and $th.MinorArab -eq 'Arial' -and $s2.Shapes.Item(1).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript -eq 'Sakkal Majalla' -and $s2.Shapes.Item(2).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript -eq 'Arial' -and (Cs $p 'AR-explicit') -eq 'Arial') "$($th.MajorArab) / $($th.MinorArab)"
    $res = Replace '(Theme Arabic heading font)|Dubai|Arabic'; [void](DropBackup $res)
    $th = ThemeFonts $p
    Record 'E5' 'default theme, heading entry again: changed a second time' ($th.MajorArab -eq 'Dubai' -and $s2.Shapes.Item(1).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript -eq 'Dubai') $th.MajorArab

    # ================================================================ a theme with no Arabic font (like the PPTK file)
    $q = $app.Presentations.Add(-1); $app.ActiveWindow.ViewType = 9
    $fs = $q.SlideMaster.Theme.ThemeFontScheme; $x0 = Join-Path $work 'empty0.xml'; $fs.Save($x0)
    $emptyXml = [regex]::Replace([IO.File]::ReadAllText($x0), '<a:font script="[^"]*" typeface="[^"]*"/>', '').Replace('typeface="Aptos Display"', 'typeface="Arial Nova Light"').Replace('typeface="Aptos"', 'typeface="Arial Nova Light"')
    $x1 = Join-Path $work 'empty1.xml'; [IO.File]::WriteAllText($x1, $emptyXml, (New-Object Text.UTF8Encoding($false))); $fs.Load($x1)
    $qs = $q.Slides.Add(1, 2); $qs.Shapes.Item(1).TextFrame2.TextRange.Text = $ar; $qs.Shapes.Item(2).TextFrame2.TextRange.Text = $ar + [char]13 + 'English'
    $q.SaveAs((Join-Path $work 'empty-theme.pptx'), 24)
    function TitleCs { $q.Slides.Item(1).Shapes.Item(1).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript }
    function BodyCs { $q.Slides.Item(1).Shapes.Item(2).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript }
    $list = [string][System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@(($prefix + 'ModulePPTKReplaceFonts.PPTK_FontsInUse')))
    $names = $list -split "`n"
    Record 'T1' 'no Arabic font in the theme: Arabic text has no font name; the list offers the two theme entries' ((TitleCs) -eq '+mj-cs' -and $names -contains '(Theme Arabic heading font)' -and $names -contains '(Theme Arabic body font)') ($names -join ', ')
    $res = Replace '(Theme Arabic heading font)|Sakkal Majalla|Arabic'; [void](DropBackup $res)
    Record 'T2' 'heading entry: Arabic titles use the new font; body and English untouched' ((TitleCs) -eq 'Sakkal Majalla' -and (BodyCs) -eq '+mn-cs' -and $q.Slides.Item(1).Shapes.Item(2).TextFrame2.TextRange.Paragraphs(2).Font.Name -eq 'Arial Nova Light' -and (Field $res 'theme') -ge 1) $res
    $res = Replace '(Theme Arabic body font)|Dubai|All'; [void](DropBackup $res)
    $bodyAfter = BodyCs
    $app.CommandBars.ExecuteMso('Undo'); Start-Sleep -Milliseconds 400
    Record 'T4' 'one Undo restores the empty body slot (heading change kept)' ($bodyAfter -eq 'Dubai' -and (BodyCs) -eq '+mn-cs' -and (TitleCs) -eq 'Sakkal Majalla') "body $bodyAfter -> $(BodyCs)"
    $res = Replace '(Theme Arabic body font)|Dubai|All'; [void](DropBackup $res)
    $qn = $q.Slides.Add(2, 2); $qn.Shapes.Item(1).TextFrame2.TextRange.Text = $ar; $qn.Shapes.Item(2).TextFrame2.TextRange.Text = $ar
    Record 'T3' 'body entry: Arabic text and new slides use the new fonts' ((BodyCs) -eq 'Dubai' -and $qn.Shapes.Item(1).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript -eq 'Sakkal Majalla' -and $qn.Shapes.Item(2).TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript -eq 'Dubai') ''
    $qn.Delete()
    # Change the heading font again; explicit text in the old heading font must not follow.
    $xb = $q.Slides.Item(1).Shapes.AddTextbox(1, 40, 300, 300, 40); $xb.Name = 'EXPLICIT-SAKKAL'; $xb.TextFrame2.TextRange.Text = $ar; $xb.TextFrame2.TextRange.Font.NameComplexScript = 'Sakkal Majalla'
    $res = Replace '(Theme Arabic heading font)|Tahoma|Arabic'; [void](DropBackup $res)
    Record 'T6' 'heading entry again (Sakkal Majalla -> Tahoma): titles change; explicit Sakkal Majalla text and body untouched' ((TitleCs) -eq 'Tahoma' -and (BodyCs) -eq 'Dubai' -and $xb.TextFrame2.TextRange.Characters(1, 5).Font.NameComplexScript -eq 'Sakkal Majalla') "title $(TitleCs) body $(BodyCs) explicit $($xb.TextFrame2.TextRange.Characters(1,5).Font.NameComplexScript)"
    $qf = $q.FullName; $q.Save(); $q.Close()
    $q = $app.Presentations.Open($qf, 0, 0, -1)
    $list = [string][System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@(($prefix + 'ModulePPTKReplaceFonts.PPTK_FontsInUse')))
    Record 'T5' 'after save and reopen: both theme Arabic fonts kept; the list shows them by name and still offers both entries' ((TitleCs) -eq 'Tahoma' -and (BodyCs) -eq 'Dubai' -and ($list -split "`n") -contains 'Dubai' -and ($list -split "`n") -contains '(Theme Arabic body font)' -and ($list -split "`n") -contains '(Theme Arabic heading font)') ($list -replace "`n", ', ')
    $q.Saved = -1; $q.Close()
    $app.ActiveWindow.View.GotoSlide(1)

    # The ribbon button and Smart Bar 62 send CMD REPLACE FONTS: it must open the PPTK window (modal, so from a
    # second process; the window is found and closed through UI Automation).
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $runner = Join-Path $work 'runner.ps1'
    "param([string]`$m, [string]`$t)`r`n`$a = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')`r`n[void][System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, `$null, `$a, [object[]]@(`$m, `$t))" | Set-Content $runner -Encoding UTF8
    $proc = Start-Process powershell -ArgumentList '-NoProfile', '-File', ('"' + $runner + '"'), ('"' + $prefix + 'ModulePPTKShortcuts.PPTK_DispatchShortcut"'), '"CMD REPLACE FONTS"' -PassThru -WindowStyle Hidden
    $pp = (Get-Process POWERPNT | Sort-Object StartTime | Select-Object -Last 1).Id
    $root = [Windows.Automation.AutomationElement]::RootElement
    $byName = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, 'Replace Fonts')
    $win = $null
    for ($i = 0; $i -lt 40 -and -not $win; $i++) {
        Start-Sleep -Milliseconds 250
        $win = @($root.FindAll([Windows.Automation.TreeScope]::Children, $byName) | Where-Object { $_.Current.ProcessId -eq $pp }) | Select-Object -First 1
        if (-not $win) { foreach ($f in @($root.FindAll([Windows.Automation.TreeScope]::Children, (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $pp))))) { $win = $f.FindFirst([Windows.Automation.TreeScope]::Children, $byName); if ($win) { break } } }
    }
    # MSForms controls are windowless and not exposed to UI Automation: the window (a UserForm, class
    # ThunderDFrame, child of the PowerPoint window) is checked, then closed like its X button.
    $isForm = $win -and $win.Current.ClassName -eq 'ThunderDFrame'
    if ($win) { try { ($win.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)).Close() } catch { } }
    $returned = $proc.WaitForExit(20000); if (-not $returned) { $proc.Kill() }
    Record 'F1' 'CMD REPLACE FONTS (ribbon button, Smart Bar 62) opens the PPTK Replace Fonts window; closing it returns' ($isForm -and $returned) "window: $(if ($win) { $win.Current.ClassName } else { 'none' })"
    $p.Saved = -1; $p.Close()
} catch {
    Record 'X' 'unexpected error' $false ($_.Exception.Message + ' at line ' + $_.InvocationInfo.ScriptLineNumber + ': ' + $_.InvocationInfo.Line.Trim())
} finally {
    try { foreach ($pr in @($app.Presentations)) { if ($pr.FullName -ne $Pptm) { $pr.Saved = -1; $pr.Close() } } } catch { }
    try { if ($addin) { $addin.Close() } } catch { }
    try { $app.Visible = 0 } catch { }
    $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    Start-Sleep 1
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
    if (-not $backupDirExisted -and (Test-Path $backupDir) -and -not (Get-ChildItem $backupDir)) { Remove-Item $backupDir -Force }
}
$failed = @($results | Where-Object { -not $_ }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
