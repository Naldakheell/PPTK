Attribute VB_Name = "ModulePPTKReplaceFonts"
Option Explicit

' PPTK Replace Fonts (Phase B): replaces one font with another across the whole presentation, Arabic included.
' Opened by the Replace Fonts button in the PPTK ribbon Text group and Smart Bar command 62 (CMD REPLACE FONTS);
' PowerPoint's own Home > Replace Fonts is not touched.
'
' Scope: All text (Latin and Arabic), Arabic only (complex-script font) or English only (Latin font).
' Order (checked in P0, 2026-11-18):
'   1. Backup copy (PPTK_SaveBackupCopy, Documents\PPTK Font Backups), then one Undo step (StartNewUndoEntry).
'   2. Theme first, in every theme (each design's slide master, notes master, handout master):
'      Latin  - ThemeFontScheme Major/Minor Latin slot, when it is the source font.
'      Arabic - the complex-script slot and the Arabic-script entries (script "Arab", "Uigh") of the theme font
'               list. The object model does not expose that list, so the scheme is saved to %TEMP%, the matching
'               typeface attributes are changed, and the scheme is loaded back (ThemeFontScheme.Save / Load).
'      Text that follows the theme now resolves to the new font and stays linked to the theme, so new slides
'      use it too.
'   3. Text walk: every run that still uses the source font is an explicit use and is changed: slides, groups,
'      tables, SmartArt nodes, notes pages, slide masters, layouts, notes and handout masters. Only the font
'      name of the chosen script is written (Font.Name for Latin, Font.NameComplexScript for Arabic); size,
'      color, bold, italic, alignment, text and geometry are never written.
'   4. English or All only: Presentation.Fonts.Replace for Latin leftovers the walk cannot reach (charts).
'      It never changes complex-script fonts (P0), so it is not used for Arabic only.
'   5. Re-scan: runs still using the source font are reported (charts and embedded objects are counted).
' Theme Arabic fonts by role: the source list always starts with PPTK_THEME_ARABIC_HEADING and
' PPTK_THEME_ARABIC_BODY. They set the Arabic font of the theme's heading or body font, empty or not, as
' often as needed (PPTK_SetThemeArabic); no text is written then.

' A theme can define no Arabic font at all (empty complex-script slot and no "Arab" entry, like the PPTK file
' theme). Arabic text that follows it then has no font name to pick (PowerPoint reports "+mj-cs" / "+mn-cs"),
' so these entries are the way to set it (checked 2026-11-18: new slides, reopening, one Undo).
Public Const PPTK_THEME_ARABIC_HEADING As String = "(Theme Arabic heading font)"
Public Const PPTK_THEME_ARABIC_BODY As String = "(Theme Arabic body font)"

Private Const ScopeAll As Long = 0
Private Const ScopeArabic As Long = 1
Private Const ScopeEnglish As Long = 2

Private mSource As String
Private mTarget As String
Private mScope As Long
Private mReplace As Boolean
Private mRuns As Long
Private mThemeSlots As Long
Private mLeft As Long
Private mFailed As Long
Private mCharts As Long
Private mFound As Object      ' Scripting.Dictionary of font names (scan)
Private mFontTimings As String

Public Sub PPTK_ShowReplaceFonts()
    On Error GoTo NoPresentation
    If ActivePresentation Is Nothing Then GoTo NoPresentation
    On Error GoTo 0
    PPTKFontsForm.PPTK_Open
    Exit Sub
NoPresentation:
    MsgBox "Open a presentation first.", vbInformation, "PPTK"
End Sub

' Font names used in the presentation (text and theme, Latin and Arabic), sorted, one per line.
Public Function PPTK_FontsInUse() As String
    Dim names() As String
    Dim count As Long
    Dim i As Long
    Dim j As Long
    Dim swapName As String
    Dim key As Variant
    Set mFound = PPTK_NewDictionary()
    mFound.CompareMode = 1
    mReplace = False
    mScope = ScopeAll
    PPTK_WalkPresentation
    PPTK_ThemeNames
    count = mFound.count
    If count = 0 Then Exit Function
    ReDim names(1 To count)
    i = 0
    For Each key In mFound.Keys
        i = i + 1
        names(i) = CStr(key)
    Next key
    For i = 1 To count - 1
        For j = i + 1 To count
            If StrComp(names(j), names(i), vbTextCompare) < 0 Then
                swapName = names(i): names(i) = names(j): names(j) = swapName
            End If
        Next j
    Next i
    PPTK_FontsInUse = Join(names, vbLf)
End Function

' Every font the user can pick in PowerPoint, sorted, one per line (read once per PowerPoint session):
'   1. every font installed in Windows, for all users and for the current user (the registry lists Windows
'      keeps, read through WMI StdRegProv, COM; no Win32 call), however it was installed;
'   2. any font file in the Windows font folders that is missing from those lists (a safety net);
'   3. Office's downloadable (cloud) fonts that PowerPoint offers, from Office's own font catalog.
' The name of an installed font is read from the font file itself (TrueType / OpenType / collection 'name'
' table, family name ID 1, the name Windows and PowerPoint show), so it does not depend on how the registry
' entry was written: installers often record only the file name ("Ubuntu-Bold.ttf"). When a file cannot be
' read, the registry name is cleaned instead ("Arial Bold (TrueType)" -> "Arial"). Nothing is font-specific.
Public Function PPTK_InstalledFonts() As String
    Static cache As String
    Static loaded As Boolean
    Dim reg As Object
    Dim shell As Object
    Dim hiveName As String
    Dim found As Object
    Dim files As Object
    Dim hive As Variant
    Dim names As Variant
    Dim types As Variant
    Dim value As Variant
    Dim index As Long
    Dim parts() As String
    Dim part As Long
    Dim fontName As String
    Dim fontPath As String
    Dim folder As Variant
    Dim fileName As String
    Dim list() As String
    Dim count As Long
    Dim i As Long
    Dim key As Variant
    Dim started As Single
    Dim mark As Single
    Dim filesRead As Long
    Const FontsKey As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    If loaded Then PPTK_InstalledFonts = cache: Exit Function
    loaded = True
    started = Timer: mark = started
    Set found = PPTK_NewDictionary()
    found.CompareMode = 1
    Set files = PPTK_NewDictionary()
    files.CompareMode = 1
    On Error Resume Next
#If Not Mac Then
    ' 1. Registry: the files Windows lists. A registry name is kept only when its file cannot be read.
    ' WMI lists the value names (two calls); the values are read with WScript.Shell.RegRead, which is fast
    ' (one WMI GetStringValue per font took ~10 ms: 6-7 s for 700 fonts). WMI is the fallback.
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Set shell = CreateObject("WScript.Shell")
    Err.Clear
    If Not reg Is Nothing Then
        For Each hive In Array(&H80000002, &H80000001)
            names = Empty
            reg.EnumValues hive, FontsKey, names, types
            If hive = &H80000002 Then hiveName = "HKLM\" Else hiveName = "HKCU\"
            If IsArray(names) Then
                For index = LBound(names) To UBound(names)
                    value = Empty
                    If Not shell Is Nothing Then value = shell.RegRead(hiveName & FontsKey & "\" & CStr(names(index)))
                    If Err.Number <> 0 Or IsEmpty(value) Then
                        Err.Clear
                        reg.GetStringValue hive, FontsKey, CStr(names(index)), value
                    End If
                    fontPath = PPTK_FontPath(CStr(value))
                    If Len(fontPath) > 0 And PPTK_IsFontFile(fontPath) And Len(Dir$(fontPath)) > 0 Then
                        If Not files.Exists(fontPath) Then files.Add fontPath, True
                    Else
                        parts = Split(PPTK_StripParens(CStr(names(index))), " & ")
                        For part = LBound(parts) To UBound(parts)
                            fontName = PPTK_FamilyName(parts(part))
                            If Len(fontName) > 0 Then If Not found.Exists(fontName) Then found.Add fontName, True
                        Next part
                    End If
                    Err.Clear
                Next index
            End If
        Next hive
    End If
#End If
    mFontTimings = "registry=" & Format$(Timer - mark, "0.00"): mark = Timer
    ' 2. Font folders (machine and current user).
#If Mac Then
    For Each folder In PPTK_MacFontFolders()
#Else
    For Each folder In Array(Environ$("WINDIR") & "\Fonts\", Environ$("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\")
#End If
        fileName = Dir$(folder & "*.*")
        Do While Len(fileName) > 0
            If PPTK_IsFontFile(fileName) Then If Not files.Exists(folder & fileName) Then files.Add folder & fileName, True
            fileName = Dir$()
        Loop
        Err.Clear
    Next folder
    mFontTimings = mFontTimings & ";folders=" & Format$(Timer - mark, "0.00"): mark = Timer
    ' Names from the files. DoEvents now and then keeps PowerPoint and the open window responsive.
    For Each key In files.Keys
        PPTK_AddFileFamilies CStr(key), found
        filesRead = filesRead + 1
        If filesRead Mod 25 = 0 Then DoEvents
    Next key
    mFontTimings = mFontTimings & ";files=" & Format$(Timer - mark, "0.00") & " (" & files.count & " files)": mark = Timer
    ' 3. Office's downloadable fonts.
    PPTK_AddCloudFonts found
    mFontTimings = mFontTimings & ";cloud=" & Format$(Timer - mark, "0.00"): mark = Timer
    On Error GoTo 0
    count = found.count
    If count = 0 Then GoTo Done
    ReDim list(1 To count)
    i = 0
    For Each key In found.Keys
        i = i + 1
        list(i) = CStr(key)
    Next key
    PPTK_SortNames list, 1, count
    cache = Join(list, vbLf)
    mFontTimings = mFontTimings & ";sort=" & Format$(Timer - mark, "0.00") & ";total=" & Format$(Timer - started, "0.00") & ";names=" & count
Done:
    PPTK_InstalledFonts = cache
End Function

' Seconds spent in each step of the last PPTK_InstalledFonts build (for tests and reports).
Public Function PPTK_InstalledFontsTiming() As String
    PPTK_InstalledFontsTiming = mFontTimings
End Function

' Quicksort, case-insensitive (the list has over a thousand names).
Private Sub PPTK_SortNames(ByRef list() As String, ByVal low As Long, ByVal high As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim swapName As String
    i = low
    j = high
    pivot = list((low + high) \ 2)
    Do While i <= j
        Do While StrComp(list(i), pivot, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(list(j), pivot, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            swapName = list(i): list(i) = list(j): list(j) = swapName
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then PPTK_SortNames list, low, j
    If i < high Then PPTK_SortNames list, i, high
End Sub

' A registry font value is a full path or a file name in the Windows font folder.
Private Function PPTK_FontPath(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) = 0 Then Exit Function
    If InStr(value, ":") > 0 Or Left$(value, 2) = "\\" Then
        PPTK_FontPath = value
    Else
        PPTK_FontPath = Environ$("WINDIR") & "\Fonts\" & value
    End If
End Function

Private Function PPTK_IsFontFile(ByVal path As String) As Boolean
    Dim ext As String
    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    PPTK_IsFontFile = (ext = "ttf" Or ext = "otf" Or ext = "ttc" Or ext = "otc")
End Function

' Adds the family name of every font in a TrueType / OpenType file or collection.
Private Sub PPTK_AddFileFamilies(ByVal path As String, ByVal found As Object)
    Dim fileNo As Integer
    Dim fontCount As Long
    Dim index As Long
    Dim fontName As String
    fileNo = FreeFile
    On Error GoTo NotOpened
    Open path For Binary Access Read Shared As #fileNo
    On Error GoTo Done
    If LOF(fileNo) < 12 Then GoTo Done
    If PPTK_Tag(fileNo, 0) = "ttcf" Then
        fontCount = PPTK_U32(fileNo, 8)
        If fontCount > 64 Then fontCount = 64
        For index = 0 To fontCount - 1
            PPTK_AddFontNames fileNo, PPTK_U32(fileNo, 12 + index * 4), found
        Next index
    Else
        PPTK_AddFontNames fileNo, 0, found
    End If
Done:
    Close #fileNo
NotOpened:
    Err.Clear
End Sub

' Adds the names one font in a file shows in Windows and PowerPoint: its family name (name ID 1) and, for a
' variable font, one name per named instance ("Cascadia Code Light", "Bahnschrift SemiBold SemiCondensed"),
' built as Windows builds it: family + instance name without Regular / Bold / Italic, cut to 31 characters.
Private Sub PPTK_AddFontNames(ByVal fileNo As Integer, ByVal base As Long, ByVal found As Object)
    Dim tableCount As Long
    Dim index As Long
    Dim record As Long
    Dim nameTable As Long
    Dim nameSize As Long
    Dim fvarTable As Long
    Dim fvarSize As Long
    Dim t() As Byte
    Dim v() As Byte
    Dim family As String
    Dim baseFamily As String
    Dim instanceName As String
    Dim fontName As String
    Dim axisCount As Long
    Dim axisSize As Long
    Dim instanceCount As Long
    Dim instanceSize As Long
    Dim at As Long
    On Error GoTo Done
    tableCount = PPTK_U16(fileNo, base + 4)
    If tableCount < 1 Or tableCount > 200 Then Exit Sub
    For index = 0 To tableCount - 1
        record = base + 12 + index * 16
        Select Case PPTK_Tag(fileNo, record)
            Case "name": nameTable = PPTK_U32(fileNo, record + 8): nameSize = PPTK_U32(fileNo, record + 12)
            Case "fvar": fvarTable = PPTK_U32(fileNo, record + 8): fvarSize = PPTK_U32(fileNo, record + 12)
        End Select
    Next index
    If nameTable = 0 Or nameSize < 6 Then Exit Sub
    If nameSize > 262144 Then nameSize = 262144
    ' The whole name table in memory: one file read per font.
    ReDim t(0 To nameSize - 1)
    Get #fileNo, nameTable + 1, t
    family = PPTK_NameText(t, nameSize, 1)
    If Len(family) > 0 Then If Not found.Exists(family) Then found.Add family, True
    If fvarTable = 0 Or fvarSize < 16 Or fvarSize > 65536 Then Exit Sub
    ReDim v(0 To fvarSize - 1)
    Get #fileNo, fvarTable + 1, v
    ' Instance names belong to the typographic family (name ID 16) when there is one.
    baseFamily = PPTK_NameText(t, nameSize, 16)
    If Len(baseFamily) = 0 Then baseFamily = family
    axisCount = CLng(v(8)) * 256 + v(9)
    axisSize = CLng(v(10)) * 256 + v(11)
    instanceCount = CLng(v(12)) * 256 + v(13)
    instanceSize = CLng(v(14)) * 256 + v(15)
    at = (CLng(v(4)) * 256 + v(5)) + axisCount * axisSize
    For index = 1 To instanceCount
        If at + 1 > fvarSize - 1 Then Exit For
        instanceName = PPTK_StyleFreeName(PPTK_NameText(t, nameSize, CLng(v(at)) * 256 + v(at + 1)))
        fontName = Trim$(baseFamily & " " & instanceName)
        If Len(fontName) > 31 Then fontName = Trim$(Left$(fontName, 31))
        If Len(fontName) > 0 Then If Not found.Exists(fontName) Then found.Add fontName, True
        at = at + instanceSize
    Next index
Done:
End Sub

' An instance name without the four style words Windows keeps as styles of one family.
Private Function PPTK_StyleFreeName(ByVal text As String) As String
    Dim word As Variant
    text = " " & text & " "
    For Each word In Array(" Regular ", " Bold ", " Italic ", " Oblique ")
        text = Replace(text, word, " ", , , vbTextCompare)
        text = Replace(text, word, " ", , , vbTextCompare)
    Next word
    PPTK_StyleFreeName = Trim$(text)
End Function

' A name from a name table held in memory: the Windows English record when there is one, else any Windows
' record, else the Macintosh record.
Private Function PPTK_NameText(ByRef t() As Byte, ByVal nameSize As Long, ByVal nameId As Long) As String
    Dim index As Long
    Dim record As Long
    Dim nameCount As Long
    Dim stringBase As Long
    Dim platform As Long
    Dim language As Long
    Dim length As Long
    Dim start As Long
    Dim score As Long
    Dim bestScore As Long
    Dim bestText As String
    Dim text As String
    Dim k As Long
    On Error GoTo Done
    nameCount = CLng(t(2)) * 256 + t(3)
    stringBase = CLng(t(4)) * 256 + t(5)
    For index = 0 To nameCount - 1
        record = 6 + index * 12
        If record + 11 > nameSize - 1 Then Exit For
        If CLng(t(record + 6)) * 256 + t(record + 7) = nameId Then
            platform = CLng(t(record)) * 256 + t(record + 1)
            language = CLng(t(record + 4)) * 256 + t(record + 5)
#If Mac Then
            ' PowerPoint for Mac lists a font by its English name. Apple fonts such as Geeza Pro have
            ' Windows records only in other languages, so the English Macintosh record ranks above them.
            If platform = 3 And language = &H409 Then
                score = 4
            ElseIf platform = 1 And language = 0 Then
                score = 3
            ElseIf platform = 3 Then
                score = 2
            ElseIf platform = 1 Then
                score = 1
            Else
                score = 0
            End If
#Else
            If platform = 3 And language = &H409 Then
                score = 3
            ElseIf platform = 3 Then
                score = 2
            ElseIf platform = 1 Then
                score = 1
            Else
                score = 0
            End If
#End If
            If score > bestScore Then
                length = CLng(t(record + 8)) * 256 + t(record + 9)
                start = stringBase + CLng(t(record + 10)) * 256 + t(record + 11)
                If length > 0 And start + length - 1 <= nameSize - 1 Then
                    text = ""
                    If platform = 3 Then
                        For k = start To start + length - 2 Step 2
                            text = text & ChrW$(CLng(t(k)) * 256 + t(k + 1))
                        Next k
                    Else
                        For k = start To start + length - 1
                            text = text & Chr$(t(k))
                        Next k
                    End If
                    text = Trim$(Replace(text, vbNullChar, ""))
                    If Len(text) > 0 Then bestScore = score: bestText = text
                End If
            End If
        End If
    Next index
    PPTK_NameText = bestText
Done:
End Function

Private Function PPTK_Tag(ByVal fileNo As Integer, ByVal at As Long) As String
    Dim bytes(0 To 3) As Byte
    Get #fileNo, at + 1, bytes
    PPTK_Tag = Chr$(bytes(0)) & Chr$(bytes(1)) & Chr$(bytes(2)) & Chr$(bytes(3))
End Function

Private Function PPTK_U16(ByVal fileNo As Integer, ByVal at As Long) As Long
    Dim bytes(0 To 1) As Byte
    Get #fileNo, at + 1, bytes
    PPTK_U16 = CLng(bytes(0)) * 256 + bytes(1)
End Function

' Offsets in font files stay far below 2 GB, so a Long is enough.
Private Function PPTK_U32(ByVal fileNo As Integer, ByVal at As Long) As Long
    Dim bytes(0 To 3) As Byte
    Get #fileNo, at + 1, bytes
    If bytes(0) > 127 Then Exit Function
    PPTK_U32 = ((CLng(bytes(0)) * 256 + bytes(1)) * 256 + bytes(2)) * 256 + bytes(3)
End Function

' Office's downloadable fonts: the "gn" (font name) entries of every Office font catalog under
' %LOCALAPPDATA%\Microsoft\FontCache\<version>\Catalog\ListAll.Json.
Private Sub PPTK_AddCloudFonts(ByVal found As Object)
    Dim root As String
    Dim version As String
    Dim versions As Collection
    Dim item As Variant
    Dim path As String
    Dim json As String
    Dim at As Long
    Dim stopAt As Long
    Dim fontName As String
    Const Marker As String = """gn"":"""
    On Error Resume Next
#If Mac Then
    root = PPTK_MacHome() & "/Library/Group Containers/UBF8T346G9.Office/FontCache/"
#Else
    root = Environ$("LOCALAPPDATA") & "\Microsoft\FontCache\"
#End If
    Set versions = New Collection
    version = Dir$(root & "*", vbDirectory)
    Do While Len(version) > 0
        If version <> "." And version <> ".." Then versions.Add version
        version = Dir$()
    Loop
    For Each item In versions
#If Mac Then
        path = root & item & "/Catalog/ListAll_hier.Json"
#Else
        path = root & item & "\Catalog\ListAll.Json"
#End If
        If Len(Dir$(path)) > 0 Then
            json = PPTK_ReadUtf8(path)
            at = InStr(1, json, Marker)
            Do While at > 0
                stopAt = InStr(at + Len(Marker), json, """")
                If stopAt = 0 Then Exit Do
                fontName = Trim$(Mid$(json, at + Len(Marker), stopAt - at - Len(Marker)))
                If Len(fontName) > 0 Then If Not found.Exists(fontName) Then found.Add fontName, True
                at = InStr(stopAt, json, Marker)
            Loop
        End If
        Err.Clear
    Next item
End Sub
Private Function PPTK_StripParens(ByVal text As String) As String
    Dim at As Long
    text = Trim$(text)
    at = InStrRev(text, " (")
    If at > 0 And Right$(text, 1) = ")" Then text = Left$(text, at - 1)
    PPTK_StripParens = Trim$(text)
End Function

Private Function PPTK_FamilyName(ByVal text As String) As String
    Dim changed As Boolean
    Dim word As Variant
    text = Trim$(text)
    If Len(text) = 0 Or Left$(text, 1) = "@" Then Exit Function
    If text Like "*.[Tt][Tt][FfCc]" Or text Like "*.[Oo][Tt][Ff]" Or text Like "*.[Ff][Oo][Nn]" Then Exit Function
    Do
        changed = False
        For Each word In Array(" Bold", " Italic", " Oblique", " Regular")
            If Len(text) > Len(word) Then
                If StrComp(Right$(text, Len(word)), word, vbTextCompare) = 0 Then
                    text = Trim$(Left$(text, Len(text) - Len(word)))
                    changed = True
                End If
            End If
        Next word
    Loop While changed
    PPTK_FamilyName = text
End Function

' spec: "source|target|scope[|force]" with scope All, Arabic or English. "force" continues without a backup.
' Returns "backup=<path>;theme=<n>;runs=<n>;left=<n>;charts=<n>;failed=<n>", or "error=<reason>",
' or "nobackup" when the backup copy could not be saved (nothing was changed).
Public Function PPTK_ReplaceFontsRun(ByVal spec As String) As String
    Dim parts() As String
    Dim backup As String
    Dim force As Boolean

    parts = Split(spec, "|")
    If UBound(parts) < 2 Then PPTK_ReplaceFontsRun = "error=bad request": Exit Function
    mSource = Trim$(parts(0))
    mTarget = Trim$(parts(1))
    Select Case LCase$(Trim$(parts(2)))
        Case "arabic": mScope = ScopeArabic
        Case "english": mScope = ScopeEnglish
        Case Else: mScope = ScopeAll
    End Select
    If UBound(parts) >= 3 Then force = (LCase$(Trim$(parts(3))) = "force")
    If Len(mSource) = 0 Or Len(mTarget) = 0 Then PPTK_ReplaceFontsRun = "error=choose both fonts": Exit Function
    If StrComp(mSource, mTarget, vbTextCompare) = 0 Then PPTK_ReplaceFontsRun = "error=the fonts are the same": Exit Function

    backup = PPTK_SaveBackupCopy(ActivePresentation, "PPTK Font Backups", "before font replace")
    If Len(backup) = 0 And Not force Then PPTK_ReplaceFontsRun = "nobackup": Exit Function

    mRuns = 0: mThemeSlots = 0: mLeft = 0: mFailed = 0: mCharts = 0
    On Error Resume Next
    PPTK_NewUndoEntry  ' Mac: the early-bound call does not compile on PowerPoint for Mac (see ModulePPTKShortcuts)
    On Error GoTo 0

    PPTK_ReplaceThemes
    If PPTK_IsThemeArabicEntry(mSource) Then
        ' Only the empty theme slot changes; text that follows it now resolves to the new font.
        PPTK_ReplaceFontsRun = "backup=" & backup & ";theme=" & mThemeSlots & ";runs=0;left=0;charts=0;failed=" & mFailed
        Exit Function
    End If
    mReplace = True
    PPTK_WalkPresentation
    If mScope <> ScopeArabic Then
        On Error Resume Next
        ActivePresentation.Fonts.Replace mSource, mTarget
        Err.Clear
        On Error GoTo 0
    End If

    ' Re-scan: what still uses the source font in the chosen scope.
    mReplace = False
    Set mFound = Nothing
    mCharts = 0
    PPTK_WalkPresentation
    PPTK_ReplaceFontsRun = "backup=" & backup & ";theme=" & mThemeSlots & ";runs=" & mRuns & _
                           ";left=" & mLeft & ";charts=" & mCharts & ";failed=" & mFailed
End Function

' ---------------------------------------------------------------- themes

Private Sub PPTK_ReplaceThemes()
    Dim index As Long
    On Error Resume Next
    For index = 1 To ActivePresentation.Designs.count
        PPTK_ReplaceTheme ActivePresentation.Designs(index).SlideMaster.Theme
    Next index
    PPTK_ReplaceTheme ActivePresentation.NotesMaster.Theme
    PPTK_ReplaceTheme ActivePresentation.HandoutMaster.Theme
    Err.Clear
End Sub

Private Sub PPTK_ReplaceTheme(ByVal theme As Object)
    Dim scheme As Object
    Dim path As String
    Dim xml As String
    Dim changed As String
    If theme Is Nothing Then Exit Sub
    On Error GoTo Failed
    Set scheme = theme.ThemeFontScheme
    If PPTK_IsThemeArabicEntry(mSource) Then
        PPTK_SetThemeArabic scheme, (mSource = PPTK_THEME_ARABIC_HEADING)
        Exit Sub
    End If
    ' Arabic first: Load replaces the whole font scheme, so the Latin change comes after it.
    If mScope <> ScopeEnglish Then
        path = PPTK_TempFolder() & "PPTK-fontscheme-" & Format$(Now, "hhnnss") & "-" & CLng(Rnd * 100000) & ".xml"
        scheme.Save path
        xml = PPTK_ReadUtf8(path)
        changed = xml
        changed = PPTK_SwapTypeface(changed, "<a:cs typeface=""")
        changed = PPTK_SwapTypeface(changed, "<a:font script=""Arab"" typeface=""")
        changed = PPTK_SwapTypeface(changed, "<a:font script=""Uigh"" typeface=""")
        If changed <> xml Then
            PPTK_WriteUtf8 path, changed
            scheme.Load path
        End If
        PPTK_DeleteFile path
    End If
    If mScope <> ScopeArabic Then
        If StrComp(scheme.MajorFont.Item(1).Name, mSource, vbTextCompare) = 0 Then scheme.MajorFont.Item(1).Name = mTarget: mThemeSlots = mThemeSlots + 1
        If StrComp(scheme.MinorFont.Item(1).Name, mSource, vbTextCompare) = 0 Then scheme.MinorFont.Item(1).Name = mTarget: mThemeSlots = mThemeSlots + 1
    End If
    Exit Sub
Failed:
    mFailed = mFailed + 1
    PPTK_DeleteFile path
    Err.Clear
End Sub

' Replaces every '<prefix>source"' with '<prefix>target"' (exact font name, case-insensitive) and counts them.
Private Function PPTK_SwapTypeface(ByVal xml As String, ByVal prefix As String) As String
    Dim findText As String
    Dim at As Long
    findText = prefix & PPTK_XmlEscape(mSource) & """"
    at = InStr(1, xml, findText, vbTextCompare)
    Do While at > 0
        xml = Left$(xml, at - 1) & prefix & PPTK_XmlEscape(mTarget) & """" & Mid$(xml, at + Len(findText))
        mThemeSlots = mThemeSlots + 1
        at = InStr(at + Len(prefix) + Len(PPTK_XmlEscape(mTarget)) + 1, xml, findText, vbTextCompare)
    Loop
    PPTK_SwapTypeface = xml
End Function

' Theme font names for the font list (Latin slots, complex-script slot, Arabic script entries).
Private Sub PPTK_ThemeNames()
    Dim index As Long
    Dim scheme As Object
    Dim path As String
    Dim xml As String
    On Error Resume Next
    For index = 1 To ActivePresentation.Designs.count
        Set scheme = ActivePresentation.Designs(index).SlideMaster.Theme.ThemeFontScheme
        PPTK_Note scheme.MajorFont.Item(1).Name
        PPTK_Note scheme.MinorFont.Item(1).Name
        path = PPTK_TempFolder() & "PPTK-fontscheme-scan-" & index & ".xml"
        scheme.Save path
        xml = PPTK_ReadUtf8(path)
        PPTK_DeleteFile path
        PPTK_NoteAttributes xml, "<a:cs typeface="""
        PPTK_NoteAttributes xml, "<a:font script=""Arab"" typeface="""
    Next index
    If ActivePresentation.Designs.count > 0 Then
        PPTK_Note PPTK_THEME_ARABIC_HEADING
        PPTK_Note PPTK_THEME_ARABIC_BODY
    End If
    Err.Clear
End Sub

Private Function PPTK_IsThemeArabicEntry(ByVal fontName As String) As Boolean
    PPTK_IsThemeArabicEntry = (fontName = PPTK_THEME_ARABIC_HEADING Or fontName = PPTK_THEME_ARABIC_BODY)
End Function

' Sets the Arabic font of the theme's heading (major) or body (minor) font, whatever it is now (empty or a
' name). A theme can name the Arabic font in two places: the script="Arab" entry (PowerPoint's own themes)
' and the complex-script slot. The entry wins when both exist (P0), so the entry is changed through the
' saved scheme (Save / edit / Load) when there is one, then the slot through the object model.
' Only this one theme font changes; text with an explicit Arabic font is not touched.
Private Sub PPTK_SetThemeArabic(ByVal scheme As Object, ByVal major As Boolean)
    Dim path As String
    Dim xml As String
    Dim tagName As String
    Dim startAt As Long
    Dim stopAt As Long
    Dim part As String
    Dim newPart As String
    Dim prefix As String
    Dim at As Long
    Dim valueEnd As Long
    On Error GoTo Clean
    If major Then tagName = "a:majorFont" Else tagName = "a:minorFont"
    path = PPTK_TempFolder() & "PPTK-fontscheme-ar-" & CLng(Rnd * 100000) & ".xml"
    scheme.Save path
    xml = PPTK_ReadUtf8(path)
    startAt = InStr(1, xml, "<" & tagName & ">", vbTextCompare)
    stopAt = InStr(1, xml, "</" & tagName & ">", vbTextCompare)
    If startAt > 0 And stopAt > startAt Then
        part = Mid$(xml, startAt, stopAt - startAt)
        prefix = "<a:font script=""Arab"" typeface="""
        at = InStr(1, part, prefix, vbTextCompare)
        If at > 0 Then
            valueEnd = InStr(at + Len(prefix), part, """")
            If valueEnd > 0 Then
                newPart = Left$(part, at + Len(prefix) - 1) & PPTK_XmlEscape(mTarget) & Mid$(part, valueEnd)
                PPTK_WriteUtf8 path, Left$(xml, startAt - 1) & newPart & Mid$(xml, stopAt)
                scheme.Load path
            End If
        End If
    End If
    PPTK_DeleteFile path
    If major Then scheme.MajorFont.Item(2).Name = mTarget Else scheme.MinorFont.Item(2).Name = mTarget
    mThemeSlots = mThemeSlots + 1
    Exit Sub
Clean:
    mFailed = mFailed + 1
    PPTK_DeleteFile path
    Err.Clear
End Sub

Private Sub PPTK_NoteAttributes(ByVal xml As String, ByVal prefix As String)
    Dim at As Long
    Dim stopAt As Long
    at = InStr(1, xml, prefix, vbTextCompare)
    Do While at > 0
        stopAt = InStr(at + Len(prefix), xml, """")
        If stopAt = 0 Then Exit Do
        PPTK_Note PPTK_XmlUnescape(Mid$(xml, at + Len(prefix), stopAt - at - Len(prefix)))
        at = InStr(stopAt, xml, prefix, vbTextCompare)
    Loop
End Sub

' ---------------------------------------------------------------- text walk

Private Sub PPTK_WalkPresentation()
    Dim sld As Slide
    Dim index As Long
    Dim layoutIndex As Long
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        PPTK_WalkShapes sld.Shapes
        If sld.HasNotesPage Then PPTK_WalkShapes sld.NotesPage.Shapes
    Next sld
    For index = 1 To ActivePresentation.Designs.count
        With ActivePresentation.Designs(index).SlideMaster
            PPTK_WalkShapes .Shapes
            For layoutIndex = 1 To .CustomLayouts.count
                PPTK_WalkShapes .CustomLayouts(layoutIndex).Shapes
            Next layoutIndex
        End With
    Next index
    PPTK_WalkShapes ActivePresentation.NotesMaster.Shapes
    PPTK_WalkShapes ActivePresentation.HandoutMaster.Shapes
    Err.Clear
End Sub

Private Sub PPTK_WalkShapes(ByVal shapes As Object)
    Dim shp As Shape
    If shapes Is Nothing Then Exit Sub
    For Each shp In shapes
        PPTK_WalkShape shp
    Next shp
End Sub

Private Sub PPTK_WalkShape(ByVal shp As Shape)
    Dim index As Long
    Dim row As Long
    Dim col As Long
    On Error GoTo Skip
    If shp.Type = msoGroup Then
        For index = 1 To shp.GroupItems.count
            PPTK_WalkShape shp.GroupItems(index)
        Next index
        Exit Sub
    End If
    If shp.HasTable Then
        For row = 1 To shp.Table.Rows.count
            For col = 1 To shp.Table.Columns.count
                PPTK_WalkRange shp.Table.Cell(row, col).Shape.TextFrame2.TextRange
            Next col
        Next row
        Exit Sub
    End If
    If shp.HasChart Then
        mCharts = mCharts + 1
        Exit Sub
    End If
    If shp.HasSmartArt Then
        For index = 1 To shp.SmartArt.AllNodes.count
            PPTK_WalkRange shp.SmartArt.AllNodes(index).TextFrame2.TextRange
        Next index
        Exit Sub
    End If
    If shp.HasTextFrame Then PPTK_WalkRange shp.TextFrame2.TextRange
    Exit Sub
Skip:
    Err.Clear
End Sub

Private Sub PPTK_WalkRange(ByVal rng As Object)
    Dim runs As Object
    Dim index As Long
    Dim run As Object
    On Error GoTo Skip
    If rng Is Nothing Then Exit Sub
    If Len(rng.Text) = 0 Then
        ' An empty placeholder still carries a font (the prompt text and what the user types next).
        PPTK_WalkRun rng
        Exit Sub
    End If
    Set runs = rng.Runs
    For index = 1 To runs.count
        Set run = rng.Runs(index)
        PPTK_WalkRun run
    Next index
    Exit Sub
Skip:
    Err.Clear
End Sub

Private Sub PPTK_WalkRun(ByVal run As Object)
    Dim latinName As String
    Dim arabicName As String
    On Error Resume Next
    latinName = run.Font.Name
    arabicName = run.Font.NameComplexScript
    Err.Clear
    If Not mFound Is Nothing And Not mReplace Then
        PPTK_Note latinName
        PPTK_Note arabicName
        Exit Sub
    End If
    If mScope <> ScopeArabic And StrComp(latinName, mSource, vbTextCompare) = 0 Then
        If mReplace Then
            run.Font.Name = mTarget
            If Err.Number = 0 Then mRuns = mRuns + 1 Else mFailed = mFailed + 1
            Err.Clear
        Else
            mLeft = mLeft + 1
        End If
    End If
    If mScope <> ScopeEnglish And StrComp(arabicName, mSource, vbTextCompare) = 0 Then
        If mReplace Then
            run.Font.NameComplexScript = mTarget
            If Err.Number = 0 Then mRuns = mRuns + 1 Else mFailed = mFailed + 1
            Err.Clear
        Else
            mLeft = mLeft + 1
        End If
    End If
End Sub

Private Sub PPTK_Note(ByVal fontName As String)
    fontName = Trim$(fontName)
    If Len(fontName) = 0 Or Left$(fontName, 1) = "+" Then Exit Sub
    If Not mFound.Exists(fontName) Then mFound.Add fontName, True
End Sub

' ---------------------------------------------------------------- files (UTF-8, Arabic font names)

#If Mac Then
Private Function PPTK_ReadUtf8(ByVal path As String) As String
    ' PowerPoint for Mac: no ADODB.Stream. Binary read, then UTF-8 decoding.
    Dim fileNo As Integer
    Dim size As Long
    Dim bytes() As Byte
    fileNo = FreeFile
    Open path For Binary Access Read As #fileNo
    size = LOF(fileNo)
    If size > 0 Then
        ReDim bytes(0 To size - 1)
        Get #fileNo, 1, bytes
    End If
    Close #fileNo
    If size > 0 Then PPTK_ReadUtf8 = PPTK_Utf8Decode(bytes, size)
End Function

Private Sub PPTK_WriteUtf8(ByVal path As String, ByVal text As String)
    Dim bytes() As Byte
    Dim size As Long
    Dim fileNo As Integer
    size = PPTK_Utf8Encode(text, bytes)
    If Len(Dir$(path)) > 0 Then Kill path
    fileNo = FreeFile
    Open path For Binary Access Write As #fileNo
    If size > 0 Then Put #fileNo, 1, bytes
    Close #fileNo
End Sub

Private Function PPTK_Utf8Decode(ByRef b() As Byte, ByVal size As Long) As String
    Dim result As String
    Dim i As Long
    Dim o As Long
    Dim c As Long
    Dim cp As Long
    result = Space$(size)
    If size >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then i = 3
    End If
    Do While i < size
        c = b(i)
        If c < &H80 Then
            cp = c
            i = i + 1
        ElseIf c < &HE0 And i + 1 < size Then
            cp = ((c And &H1F) * &H40&) Or (b(i + 1) And &H3F)
            i = i + 2
        ElseIf c < &HF0 And i + 2 < size Then
            cp = ((c And &HF) * &H1000&) Or ((b(i + 1) And &H3F) * &H40&) Or (b(i + 2) And &H3F)
            i = i + 3
        ElseIf i + 3 < size Then
            cp = ((c And &H7) * &H40000) Or ((b(i + 1) And &H3F) * &H1000&) Or ((b(i + 2) And &H3F) * &H40&) Or (b(i + 3) And &H3F)
            i = i + 4
        Else
            cp = &HFFFD&
            i = i + 1
        End If
        If cp > &HFFFF& Then
            cp = cp - &H10000
            o = o + 1
            Mid$(result, o, 1) = ChrW$(&HD800& + (cp \ &H400&))
            o = o + 1
            Mid$(result, o, 1) = ChrW$(&HDC00& + (cp And &H3FF&))
        Else
            o = o + 1
            Mid$(result, o, 1) = ChrW$(cp)
        End If
    Loop
    PPTK_Utf8Decode = Left$(result, o)
End Function

Private Function PPTK_Utf8Encode(ByVal text As String, ByRef bytes() As Byte) As Long
    Dim i As Long
    Dim n As Long
    Dim cp As Long
    Dim low As Long
    If Len(text) = 0 Then Exit Function
    ReDim bytes(0 To Len(text) * 4)
    i = 1
    Do While i <= Len(text)
        cp = AscW(Mid$(text, i, 1)) And &HFFFF&
        If cp >= &HD800& And cp <= &HDBFF& And i < Len(text) Then
            low = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
            cp = &H10000 + (cp - &HD800&) * &H400& + (low - &HDC00&)
            i = i + 1
        End If
        If cp < &H80 Then
            bytes(n) = cp: n = n + 1
        ElseIf cp < &H800 Then
            bytes(n) = &HC0 Or (cp \ &H40&): bytes(n + 1) = &H80 Or (cp And &H3F): n = n + 2
        ElseIf cp < &H10000 Then
            bytes(n) = &HE0 Or (cp \ &H1000&): bytes(n + 1) = &H80 Or ((cp \ &H40&) And &H3F): bytes(n + 2) = &H80 Or (cp And &H3F): n = n + 3
        Else
            bytes(n) = &HF0 Or (cp \ &H40000): bytes(n + 1) = &H80 Or ((cp \ &H1000&) And &H3F)
            bytes(n + 2) = &H80 Or ((cp \ &H40&) And &H3F): bytes(n + 3) = &H80 Or (cp And &H3F): n = n + 4
        End If
        i = i + 1
    Loop
    ReDim Preserve bytes(0 To n - 1)
    PPTK_Utf8Encode = n
End Function
#Else
Private Function PPTK_ReadUtf8(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    PPTK_ReadUtf8 = stream.ReadText
    stream.Close
End Function

Private Sub PPTK_WriteUtf8(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile path, 2
    stream.Close
End Sub
#End If

Private Sub PPTK_DeleteFile(ByVal path As String)
    On Error Resume Next
    If Len(path) > 0 Then If Len(Dir$(path)) > 0 Then Kill path
    Err.Clear
End Sub

Private Function PPTK_XmlEscape(ByVal text As String) As String
    PPTK_XmlEscape = Replace(Replace(Replace(Replace(text, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;")
End Function

Private Function PPTK_XmlUnescape(ByVal text As String) As String
    PPTK_XmlUnescape = Replace(Replace(Replace(Replace(text, "&quot;", """"), "&gt;", ">"), "&lt;", "<"), "&amp;", "&")
End Function

' ---------------------------------------------------------------- PowerPoint for Mac adaptations

Private Function PPTK_NewDictionary() As Object
#If Mac Then
    Set PPTK_NewDictionary = New PPTKMacDict
#Else
    Set PPTK_NewDictionary = CreateObject("Scripting.Dictionary")
#End If
End Function

Private Function PPTK_TempFolder() As String
#If Mac Then
    PPTK_TempFolder = Environ$("TMPDIR")
#Else
    PPTK_TempFolder = Environ$("TEMP") & "\"
#End If
End Function

#If Mac Then
' The real home folder (Environ HOME is PowerPoint's sandbox container).
Private Function PPTK_MacHome() As String
    Dim home As String
    home = Environ$("HOME")
    If InStr(home, "/Library/Containers/") > 0 Then home = Left$(home, InStr(home, "/Library/Containers/") - 1)
    PPTK_MacHome = home
End Function

' System, local, user and Office-bundled fonts (PowerPoint for Mac ships Calibri, Aptos and others in DFonts).
Private Function PPTK_MacFontFolders() As Variant
    PPTK_MacFontFolders = Array("/System/Library/Fonts/", "/System/Library/Fonts/Supplemental/", "/Library/Fonts/", _
        PPTK_MacHome() & "/Library/Fonts/", "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/")
End Function
#End If
