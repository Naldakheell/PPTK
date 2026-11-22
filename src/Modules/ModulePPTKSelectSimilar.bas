Attribute VB_Name = "ModulePPTKSelectSimilar"
Option Explicit

' PPTK Select Similar (ribbon Tools > Select Similar, Smart Bar): selects the shapes on the current slide that
' match the first selected shape (the same reference rule as ModulePPTKAlign).
'
'   Fill color     visible solid fill with the same displayed RGB
'   Outline color  visible outline with the same displayed RGB
'   Font           any text that uses the reference's font. Arabic text is compared by its complex-script font
'                  (Font.NameComplexScript), other text by Font.Name.
'   Font color     any text with the reference's displayed text color
'   Font size      any text with the reference's font size
'   Shape type     the same kind of object (and the same AutoShape type for AutoShapes)
'   Width, Height  the same width / height (within half a point)
'   Size           the same width and height (within half a point)
'
' Value lists (the right-click Select menu, ModulePPTKSelectMenu): PPTK_SelectCollect lists the distinct values of
' one property on the current slide (colors, font sizes, object types, widths, heights, sizes, rounded to 0.1 pt),
' and PPTK_SelectByValue selects every shape with one of those values. No reference shape is needed.
'
' Colors are compared on the displayed RGB (ColorFormat.RGB), which is what PowerPoint shows for explicit RGB,
' theme colors, Lighter/Darker steps and style tints (see ModulePPTKReplaceColors). The reference's property is
' the first one found: its fill, its outline, the font or color of its first text run.
'
' Scope: the current slide only. PowerPoint can select shapes on one slide at a time. Hidden shapes are not
' selected. Groups and tables are one object each; their text (any depth, every cell) counts for Font and Font
' color. The reference stays first in the new selection, so a following Align uses it again.
' Selecting changes nothing on the slide. No Win32 or Windows-only objects.

Private Const SizeTolerance As Single = 0.5

' Reference size for Size (set by PPTK_SelectSimilar before matching).
Private mRefWidth As Single
Private mRefHeight As Single

' Value list from PPTK_SelectCollect (read by the right-click Select menu).
Private mListKind As String
Private mListValues() As String
Private mListLabels() As String
Private mListCount As Long

Public Sub PPTK_SelectSimilarFill()
    PPTK_SelectSimilar "FILL"
End Sub

Public Sub PPTK_SelectSimilarOutline()
    PPTK_SelectSimilar "OUTLINE"
End Sub

Public Sub PPTK_SelectSimilarFont()
    PPTK_SelectSimilar "FONT"
End Sub

Public Sub PPTK_SelectSimilarFontColor()
    PPTK_SelectSimilar "FONT COLOR"
End Sub

Public Sub PPTK_SelectSimilarType()
    PPTK_SelectSimilar "TYPE"
End Sub

Public Sub PPTK_SelectSimilarSize()
    PPTK_SelectSimilar "SIZE"
End Sub

Public Sub PPTK_SelectSimilarFontSize()
    PPTK_SelectSimilar "FONT SIZE"
End Sub

Public Sub PPTK_SelectSimilarWidth()
    PPTK_SelectSimilar "WIDTH"
End Sub

Public Sub PPTK_SelectSimilarHeight()
    PPTK_SelectSimilar "HEIGHT"
End Sub

' kind: FILL, OUTLINE, FONT, FONT COLOR, FONT SIZE, TYPE, WIDTH, HEIGHT, SIZE. Returns the number of shapes selected (0 when nothing to do).
Public Function PPTK_SelectSimilar(ByVal kind As String) As Long
    Dim win As Object
    Dim sld As Slide
    Dim reference As Shape
    Dim key As String
    Dim indexes() As Variant
    Dim count As Long
    Dim index As Long
    Dim shp As Shape

    On Error GoTo Clean
    Set win = Application.ActiveWindow
    If win.Selection.Type <> ppSelectionShapes Then
        MsgBox "Select one shape first; PPTK then selects the shapes that match it.", vbInformation, "PPTK"
        Exit Function
    End If
    If win.Selection.HasChildShapeRange Then Set reference = win.Selection.ChildShapeRange(1) Else Set reference = win.Selection.ShapeRange(1)
    Set sld = win.View.Slide
    mRefWidth = reference.Width
    mRefHeight = reference.Height
    key = PPTK_SimilarKey(reference, kind)
    If key = "" Then
        MsgBox "The selected shape has no " & PPTK_KindLabel(kind) & " to match.", vbInformation, "PPTK"
        Exit Function
    End If

    ' The reference first (a top-level shape, or the group that holds it), then every other match in slide order.
    ReDim indexes(1 To sld.Shapes.Count)
    For index = 1 To sld.Shapes.Count
        If sld.Shapes(index).Id = PPTK_TopLevel(reference).Id Then
            count = 1
            indexes(1) = index
        End If
    Next index
    If count = 0 Then Exit Function
    For index = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(index)
        If index <> indexes(1) And shp.Visible = msoTrue Then
            If PPTK_Matches(shp, kind, key) Then
                count = count + 1
                indexes(count) = index
            End If
        End If
    Next index
    If count = 1 Then
        MsgBox "No other shape on this slide has the same " & PPTK_KindLabel(kind) & ".", vbInformation, "PPTK"
        PPTK_SelectSimilar = 1
        Exit Function
    End If
    ReDim Preserve indexes(1 To count)
    sld.Shapes.Range(indexes).Select
    PPTK_SelectSimilar = count
Clean:
End Function

Private Function PPTK_KindLabel(ByVal kind As String) As String
    Select Case kind
        Case "FILL": PPTK_KindLabel = "fill color"
        Case "OUTLINE": PPTK_KindLabel = "outline color"
        Case "FONT": PPTK_KindLabel = "font"
        Case "FONT COLOR": PPTK_KindLabel = "font color"
        Case "TYPE": PPTK_KindLabel = "shape type"
        Case "SIZE": PPTK_KindLabel = "size"
        Case "FONT SIZE": PPTK_KindLabel = "font size"
        Case "WIDTH": PPTK_KindLabel = "width"
        Case "HEIGHT": PPTK_KindLabel = "height"
    End Select
End Function

Private Function PPTK_TopLevel(ByVal shp As Shape) As Shape
    On Error Resume Next
    Set PPTK_TopLevel = shp
    Set PPTK_TopLevel = shp.ParentGroup
    Err.Clear
End Function

' ---------------------------------------------------------------- what is compared

' The reference value as text ("" when the reference has nothing to compare).
Private Function PPTK_SimilarKey(ByVal shp As Shape, ByVal kind As String) As String
    Dim values As Collection
    Select Case kind
        Case "FILL": PPTK_SimilarKey = PPTK_FillKey(shp)
        Case "OUTLINE": PPTK_SimilarKey = PPTK_LineKey(shp)
        Case "TYPE": PPTK_SimilarKey = PPTK_TypeKey(shp)
        Case "SIZE", "WIDTH", "HEIGHT": PPTK_SimilarKey = "size"
        Case "FONT", "FONT COLOR", "FONT SIZE"
            Set values = New Collection
            PPTK_TextValues shp, kind, values
            If values.Count > 0 Then PPTK_SimilarKey = values(1)
    End Select
End Function

Private Function PPTK_Matches(ByVal shp As Shape, ByVal kind As String, ByVal key As String) As Boolean
    Dim values As Collection
    Dim value As Variant
    Select Case kind
        Case "FILL": PPTK_Matches = (PPTK_FillKey(shp) = key)
        Case "OUTLINE": PPTK_Matches = (PPTK_LineKey(shp) = key)
        Case "TYPE": PPTK_Matches = (PPTK_TypeKey(shp) = key)
        Case "SIZE": PPTK_Matches = PPTK_SameSize(shp)
        Case "WIDTH": PPTK_Matches = (Abs(shp.Width - mRefWidth) <= SizeTolerance)
        Case "HEIGHT": PPTK_Matches = (Abs(shp.Height - mRefHeight) <= SizeTolerance)
        Case "FONT", "FONT COLOR", "FONT SIZE"
            Set values = New Collection
            PPTK_TextValues shp, kind, values
            For Each value In values
                If value = key Then PPTK_Matches = True: Exit Function
            Next value
    End Select
End Function

Private Function PPTK_FillKey(ByVal shp As Shape) As String
    On Error GoTo NoKey
    If shp.Type = msoGroup Or shp.Type = msoTable Or shp.Type = msoPicture Then Exit Function
    If shp.Fill.Visible <> msoTrue Then Exit Function
    If shp.Fill.Type <> msoFillSolid Then Exit Function
    PPTK_FillKey = CStr(shp.Fill.ForeColor.RGB)
NoKey:
End Function

Private Function PPTK_LineKey(ByVal shp As Shape) As String
    On Error GoTo NoKey
    If shp.Type = msoGroup Or shp.Type = msoTable Then Exit Function
    If shp.Line.Visible <> msoTrue Then Exit Function
    PPTK_LineKey = CStr(shp.Line.ForeColor.RGB)
NoKey:
End Function

Private Function PPTK_TypeKey(ByVal shp As Shape) As String
    On Error GoTo NoKey
    PPTK_TypeKey = CStr(shp.Type)
    If shp.Type = msoAutoShape Then PPTK_TypeKey = PPTK_TypeKey & ":" & CStr(shp.AutoShapeType)
NoKey:
End Function

' Size is compared with the reference size saved by PPTK_SelectSimilar (mRefWidth, mRefHeight).
Private Function PPTK_SameSize(ByVal shp As Shape) As Boolean
    PPTK_SameSize = (Abs(shp.Width - mRefWidth) <= SizeTolerance And Abs(shp.Height - mRefHeight) <= SizeTolerance)
End Function

' ---------------------------------------------------------------- text (fonts and font colors of every run)

' Adds the font names (FONT), font sizes (FONT SIZE) or displayed RGB values (FONT COLOR) of every text run in the shape, its table
' cells and its group members, in reading order.
Private Sub PPTK_TextValues(ByVal shp As Shape, ByVal kind As String, ByVal values As Collection)
    Dim child As Shape
    Dim row As Long
    Dim col As Long
    On Error Resume Next
    If shp.Type = msoGroup Then
        For Each child In shp.GroupItems
            PPTK_TextValues child, kind, values
        Next child
    ElseIf shp.HasTable Then
        For row = 1 To shp.Table.Rows.Count
            For col = 1 To shp.Table.Columns.Count
                PPTK_RunValues shp.Table.Cell(row, col).Shape.TextFrame2.TextRange, kind, values
            Next col
        Next row
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame2.HasText Then PPTK_RunValues shp.TextFrame2.TextRange, kind, values
    End If
    Err.Clear
End Sub

Private Sub PPTK_RunValues(ByVal rng As TextRange2, ByVal kind As String, ByVal values As Collection)
    Dim run As TextRange2
    Dim index As Long
    On Error Resume Next
    For index = 1 To rng.Runs.Count
        Set run = rng.Runs(index)
        If Len(Trim$(run.Text)) > 0 Then
            Select Case kind
                Case "FONT"
                    If PPTK_HasArabic(run.Text) Then values.Add CStr(run.Font.NameComplexScript)
                    If PPTK_HasNonArabicLetters(run.Text) Then values.Add CStr(run.Font.Name)
                Case "FONT SIZE"
                    values.Add CStr(run.Font.Size)
                Case Else
                    values.Add CStr(run.Font.Fill.ForeColor.RGB)
            End Select
        End If
    Next index
    Err.Clear
End Sub

Private Function PPTK_HasArabic(ByVal text As String) As Boolean
    Dim index As Long
    Dim code As Long
    For index = 1 To Len(text)
        code = AscW(Mid$(text, index, 1)) And &HFFFF&
        If (code >= &H600& And code <= &H6FF&) Or (code >= &H750& And code <= &H77F&) Or (code >= &HFB50& And code <= &HFDFF&) Or (code >= &HFE70& And code <= &HFEFF&) Then
            PPTK_HasArabic = True
            Exit Function
        End If
    Next index
End Function

Private Function PPTK_HasNonArabicLetters(ByVal text As String) As Boolean
    Dim index As Long
    Dim ch As String
    Dim code As Long
    For index = 1 To Len(text)
        ch = Mid$(text, index, 1)
        code = AscW(ch) And &HFFFF&
        If ch Like "[A-Za-z0-9]" Or (code >= &HC0& And code < &H600&) Then
            PPTK_HasNonArabicLetters = True
            Exit Function
        End If
    Next index
End Function

' ---------------------------------------------------------------- value lists (right-click Select menu)


' The distinct values of one property on the current slide, in slide order. kind: FILL, OUTLINE, FONT COLOR,
' FONT SIZE, TYPE, WIDTH, HEIGHT, SIZE. Returns how many; read them with PPTK_SelectValueAt/PPTK_SelectLabelAt.
Public Function PPTK_SelectCollect(ByVal kind As String) As Long
    Dim sld As Slide
    Dim shp As Shape
    Dim value As Variant
    mListKind = kind
    mListCount = 0
    ReDim mListValues(1 To 1)
    ReDim mListLabels(1 To 1)
    On Error GoTo Clean
    Set sld = Application.ActiveWindow.View.Slide
    For Each shp In sld.Shapes
        If shp.Visible = msoTrue Then
            For Each value In PPTK_ValueKeysOf(shp, kind)
                If PPTK_ListIndexOf(CStr(value)) = 0 Then
                    mListCount = mListCount + 1
                    ReDim Preserve mListValues(1 To mListCount)
                    ReDim Preserve mListLabels(1 To mListCount)
                    mListValues(mListCount) = CStr(value)
                    mListLabels(mListCount) = PPTK_ValueLabel(kind, CStr(value))
                End If
            Next value
        End If
    Next shp
Clean:
    PPTK_SelectCollect = mListCount
End Function

Public Function PPTK_SelectValueAt(ByVal index As Long) As String
    If index >= 1 And index <= mListCount Then PPTK_SelectValueAt = mListValues(index)
End Function

Public Function PPTK_SelectLabelAt(ByVal index As Long) As String
    If index >= 1 And index <= mListCount Then PPTK_SelectLabelAt = mListLabels(index)
End Function

' spec: "<kind>|<value>" (a value from PPTK_SelectCollect). Selects every visible shape on the current slide that
' has it; returns how many (0: nothing selected, nothing changed).
Public Function PPTK_SelectByValue(ByVal spec As String) As Long
    Dim sld As Slide
    Dim shp As Shape
    Dim indexes() As Variant
    Dim count As Long
    Dim index As Long
    Dim kind As String
    Dim wanted As String
    Dim value As Variant
    Dim cut As Long
    cut = InStr(spec, "|")
    If cut < 2 Then Exit Function
    kind = Left$(spec, cut - 1)
    wanted = Mid$(spec, cut + 1)
    On Error GoTo Clean
    Set sld = Application.ActiveWindow.View.Slide
    If sld.Shapes.Count = 0 Then Exit Function
    ReDim indexes(1 To sld.Shapes.Count)
    For index = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(index)
        If shp.Visible = msoTrue Then
            For Each value In PPTK_ValueKeysOf(shp, kind)
                If CStr(value) = wanted Then
                    count = count + 1
                    indexes(count) = index
                    Exit For
                End If
            Next value
        End If
    Next index
    If count = 0 Then Exit Function
    ReDim Preserve indexes(1 To count)
    sld.Shapes.Range(indexes).Select
    PPTK_SelectByValue = count
Clean:
End Function

' Every visible shape on the current slide that contains text (groups and tables included).
Public Function PPTK_SelectAllText() As Long
    Dim sld As Slide
    Dim indexes() As Variant
    Dim count As Long
    Dim index As Long
    Dim runs As Collection
    On Error GoTo Clean
    Set sld = Application.ActiveWindow.View.Slide
    If sld.Shapes.Count = 0 Then Exit Function
    ReDim indexes(1 To sld.Shapes.Count)
    For index = 1 To sld.Shapes.Count
        If sld.Shapes(index).Visible = msoTrue Then
            Set runs = New Collection
            PPTK_TextValues sld.Shapes(index), "FONT SIZE", runs
            If runs.Count > 0 Then
                count = count + 1
                indexes(count) = index
            End If
        End If
    Next index
    If count = 0 Then Exit Function
    ReDim Preserve indexes(1 To count)
    sld.Shapes.Range(indexes).Select
    PPTK_SelectAllText = count
Clean:
End Function

Private Function PPTK_ListIndexOf(ByVal value As String) As Long
    Dim index As Long
    For index = 1 To mListCount
        If mListValues(index) = value Then
            PPTK_ListIndexOf = index
            Exit Function
        End If
    Next index
End Function

' The property values one shape has (several for text, which can mix colors and sizes).
Private Function PPTK_ValueKeysOf(ByVal shp As Shape, ByVal kind As String) As Collection
    Dim values As New Collection
    Dim key As String
    On Error Resume Next
    Select Case kind
        Case "FILL": key = PPTK_FillKey(shp): If key <> "" Then values.Add key
        Case "OUTLINE": key = PPTK_LineKey(shp): If key <> "" Then values.Add key
        Case "TYPE": key = PPTK_TypeKey(shp): If key <> "" Then values.Add key
        Case "FONT COLOR", "FONT SIZE": PPTK_TextValues shp, kind, values
        Case "WIDTH": values.Add PPTK_Points(shp.Width)
        Case "HEIGHT": values.Add PPTK_Points(shp.Height)
        Case "SIZE": values.Add PPTK_Points(shp.Width) & "x" & PPTK_Points(shp.Height)
    End Select
    Err.Clear
    Set PPTK_ValueKeysOf = values
End Function

Private Function PPTK_Points(ByVal value As Single) As String
    PPTK_Points = Format$(Round(value, 1), "0.0")
End Function

' Menu text for a value: sizes in points, object type names; colors as #RRGGBB (shown only as a tooltip).
Public Function PPTK_ValueLabel(ByVal kind As String, ByVal value As String) As String
    Dim parts() As String
    Select Case kind
        Case "FONT SIZE", "WIDTH", "HEIGHT": PPTK_ValueLabel = value & " pt"
        Case "SIZE": parts = Split(value, "x"): PPTK_ValueLabel = parts(0) & " " & ChrW$(&HD7) & " " & parts(1) & " pt"
        Case "TYPE": PPTK_ValueLabel = PPTK_TypeName(value)
        Case Else: PPTK_ValueLabel = PPTK_HexColor(CLng(value))
    End Select
End Function

Public Function PPTK_HexColor(ByVal rgbValue As Long) As String
    PPTK_HexColor = "#" & Right$("0" & Hex$(rgbValue And &HFF&), 2) & Right$("0" & Hex$((rgbValue \ &H100&) And &HFF&), 2) & Right$("0" & Hex$((rgbValue \ &H10000) And &HFF&), 2)
End Function

' key: PPTK_TypeKey ("<Shape.Type>" or "1:<AutoShapeType>" for AutoShapes).
Private Function PPTK_TypeName(ByVal key As String) As String
    Select Case key
        Case "1:1": PPTK_TypeName = "Rectangle"
        Case "1:5": PPTK_TypeName = "Rounded Rectangle"
        Case "1:9": PPTK_TypeName = "Oval"
        Case "1:7": PPTK_TypeName = "Triangle"
        Case "1:8": PPTK_TypeName = "Right Triangle"
        Case "1:52": PPTK_TypeName = "Chevron"
        Case "1:51": PPTK_TypeName = "Pentagon"
        Case "1:2": PPTK_TypeName = "Parallelogram"
        Case "1:33": PPTK_TypeName = "Right Arrow"
        Case "17": PPTK_TypeName = "Text Box"
        Case "13": PPTK_TypeName = "Picture"
        Case "11": PPTK_TypeName = "Linked Picture"
        Case "28": PPTK_TypeName = "Icon"
        Case "6": PPTK_TypeName = "Group"
        Case "19": PPTK_TypeName = "Table"
        Case "9": PPTK_TypeName = "Line"
        Case "14": PPTK_TypeName = "Placeholder"
        Case "3": PPTK_TypeName = "Chart"
        Case "24": PPTK_TypeName = "SmartArt"
        Case "5": PPTK_TypeName = "Freeform"
        Case Else
            If Left$(key, 2) = "1:" Then PPTK_TypeName = "Shape " & Mid$(key, 3) Else PPTK_TypeName = "Object " & key
    End Select
End Function