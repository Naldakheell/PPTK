Attribute VB_Name = "ModulePPTKSelectMenu"
Option Explicit

' PPTK right-click Select menu (Auxy layout), in PowerPoint's own context menus for the slide background
' (ContextMenuFrame) and for shapes, groups, pictures and icons (src\PPTK\Ribbon\CustomUI.xml, <contextMenus>):
'
'   Select >
'     Select All                       PowerPoint's own Select All
'     Select All Text on Slide         every shape on the slide that contains text
'     Select By Fill Color >           swatches of the fill colors on this slide
'     Select By Border Color >         swatches of the outline colors
'     Select By Font Color >           swatches of the text colors
'     Select By Font Size >            the font sizes on this slide
'     Select By Object Type >          the object types
'     Select Shapes of Same Width >    the widths
'     Select Shapes of Same Height >   the heights
'     Select Shapes of Same Size >     the sizes
'
' Choosing a value selects every shape on the current slide with that value (ModulePPTKSelectSimilar:
' PPTK_SelectCollect, PPTK_SelectByValue). Nothing needs to be selected first; nothing on the slide changes.
' Each "Select By" entry is a dynamicMenu, so its values are read when it opens. Colors are a gallery of swatches
' without text (the value is the tooltip); a new gallery id each time makes PowerPoint ask for the current
' swatches. Swatch images: a 16 x 16 BMP written to %TEMP% and loaded with stdole LoadPicture (no Win32 calls).
' Context menus are a Windows (Office 2010+) feature; on the Mac the ribbon and Smart Bar commands remain.
' Checked in PowerPoint on 2026-11-17 (S0): menu, nesting, swatches and selection.

Private mSerial As Long

Public Sub PPTK_SelectMenuAll(ByVal control As IRibbonControl)
    On Error Resume Next
    Application.CommandBars.ExecuteMso "SelectAll"
End Sub

Public Sub PPTK_SelectMenuText(ByVal control As IRibbonControl)
    PPTK_SelectAllText
End Sub

' getContent of each "Select By" dynamicMenu (its tag is the property).
Public Sub PPTK_SelectMenuContent(ByVal control As IRibbonControl, ByRef content)
    content = PPTK_SelectMenuXml(control.Tag)
End Sub

Public Function PPTK_SelectMenuXml(ByVal kind As String) As String
    Dim xml As String
    Dim count As Long
    Dim index As Long
    count = PPTK_SelectCollect(kind)
    mSerial = mSerial + 1
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"
    If count = 0 Then
        xml = xml & "<button id=""PPTKSelNone" & mSerial & """ label=""None on this slide"" enabled=""false""/>"
    ElseIf PPTK_SelectIsColor(kind) Then
        xml = xml & "<gallery id=""PPTKSelSwatch" & mSerial & """ label=""" & PPTK_XmlText(kind) & """ showLabel=""false""" & _
              " columns=""" & IIf(count < 8, count, 8) & """ itemWidth=""16"" itemHeight=""16"" showItemLabel=""false""" & _
              " tag=""" & PPTK_XmlText(kind) & """ getItemCount=""PPTK_SelectSwatchCount"" getItemID=""PPTK_SelectSwatchID""" & _
              " getItemImage=""PPTK_SelectSwatchImage"" getItemScreentip=""PPTK_SelectSwatchTip"" onAction=""PPTK_SelectSwatchPick""/>"
    Else
        For index = 1 To count
            xml = xml & "<button id=""PPTKSelValue" & mSerial & "_" & index & """ label=""" & PPTK_XmlText(PPTK_SelectLabelAt(index)) & """" & _
                  " tag=""" & PPTK_XmlText(kind & "|" & PPTK_SelectValueAt(index)) & """ onAction=""PPTK_SelectMenuPick""/>"
        Next index
    End If
    PPTK_SelectMenuXml = xml & "</menu>"
End Function

Public Sub PPTK_SelectMenuPick(ByVal control As IRibbonControl)
    PPTK_SelectByValue control.Tag
End Sub

Private Function PPTK_SelectIsColor(ByVal kind As String) As Boolean
    PPTK_SelectIsColor = (kind = "FILL" Or kind = "OUTLINE" Or kind = "FONT COLOR")
End Function

' ---------------------------------------------------------------- swatch gallery (values from the last PPTK_SelectCollect)

Public Sub PPTK_SelectSwatchCount(ByVal control As IRibbonControl, ByRef count)
    count = PPTK_SelectCollect(control.Tag)
End Sub

Public Sub PPTK_SelectSwatchID(ByVal control As IRibbonControl, ByVal index As Integer, ByRef id)
    id = "PPTKSwatch" & index
End Sub

Public Sub PPTK_SelectSwatchImage(ByVal control As IRibbonControl, ByVal index As Integer, ByRef image)
    On Error Resume Next
    Set image = LoadPicture(PPTK_SwatchFile(CLng(PPTK_SelectValueAt(index + 1))))
End Sub

Public Sub PPTK_SelectSwatchTip(ByVal control As IRibbonControl, ByVal index As Integer, ByRef screentip)
    screentip = PPTK_SelectLabelAt(index + 1)
End Sub

Public Sub PPTK_SelectSwatchPick(ByVal control As IRibbonControl, ByVal selectedId As String, ByVal selectedIndex As Integer)
    PPTK_SelectByValue control.Tag & "|" & PPTK_SelectValueAt(selectedIndex + 1)
End Sub

' A 16 x 16 24-bit BMP of one color with a thin gray border, written once to %TEMP%. Returns its path.
Public Function PPTK_SwatchFile(ByVal rgbValue As Long) As String
    Const side As Long = 16
    Dim path As String
    Dim fileNo As Integer
    Dim data() As Byte
    Dim at As Long
    Dim x As Long
    Dim y As Long
    Dim dataSize As Long
    path = Environ$("TEMP") & "\PPTK-swatch-" & Right$("00000" & Hex$(rgbValue), 6) & ".bmp"
    If Len(Dir$(path)) > 0 Then PPTK_SwatchFile = path: Exit Function
    dataSize = side * side * 3                   ' each row is 48 bytes, already a multiple of 4
    ReDim data(0 To 53 + dataSize)
    data(0) = 66: data(1) = 77                   ' "BM"
    PPTK_PutLong data, 2, 54 + dataSize
    PPTK_PutLong data, 10, 54
    PPTK_PutLong data, 14, 40
    PPTK_PutLong data, 18, side
    PPTK_PutLong data, 22, side
    data(26) = 1
    data(28) = 24
    PPTK_PutLong data, 34, dataSize
    at = 54
    For y = 0 To side - 1
        For x = 0 To side - 1
            If x = 0 Or y = 0 Or x = side - 1 Or y = side - 1 Then
                data(at) = 160: data(at + 1) = 160: data(at + 2) = 160
            Else
                data(at) = (rgbValue \ &H10000) And &HFF&
                data(at + 1) = (rgbValue \ &H100&) And &HFF&
                data(at + 2) = rgbValue And &HFF&
            End If
            at = at + 3
        Next x
    Next y
    fileNo = FreeFile
    Open path For Binary Access Write As #fileNo
    Put #fileNo, , data
    Close #fileNo
    PPTK_SwatchFile = path
End Function

Private Sub PPTK_PutLong(ByRef data() As Byte, ByVal at As Long, ByVal value As Long)
    data(at) = value And &HFF&
    data(at + 1) = (value \ &H100&) And &HFF&
    data(at + 2) = (value \ &H10000) And &HFF&
    data(at + 3) = (value \ &H1000000) And &HFF&
End Sub

Private Function PPTK_XmlText(ByVal text As String) As String
    PPTK_XmlText = Replace(Replace(Replace(Replace(text, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;")
End Function
