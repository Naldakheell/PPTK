Attribute VB_Name = "ModuleObjectsSizeAndPosition"
'MIT License

'Copyright (c) 2021 - 2026 iappyx

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Sub ObjectsSizeToTallest(Optional ByVal control As IRibbonControl)
    Set MyDocument = Application.ActiveWindow
    Dim height      As Single
    Dim Tallest     As Single
    Dim SlideShape  As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Tallest = GetRealHeight(MyDocument.Selection.ChildShapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            height = GetRealHeight(SlideShape)
            If height > Tallest Then Tallest = height
        Next
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            SetRealHeight SlideShape, Tallest
        Next SlideShape
        
    Else
        Tallest = GetRealHeight(MyDocument.Selection.ShapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            height = GetRealHeight(SlideShape)
            If height > Tallest Then Tallest = height
        Next
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            SetRealHeight SlideShape, Tallest
        Next SlideShape
        
    End If
    
End Sub

Sub ObjectsSizeToShortest(Optional ByVal control As IRibbonControl)
    Set MyDocument = Application.ActiveWindow
    Dim height      As Single
    Dim Shortest    As Single
    Dim SlideShape  As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Shortest = GetRealHeight(MyDocument.Selection.ChildShapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            height = GetRealHeight(SlideShape)
            If height < Shortest Then Shortest = height
        Next
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            SetRealHeight SlideShape, Shortest
        Next SlideShape
        
    Else
        
        Shortest = GetRealHeight(MyDocument.Selection.ShapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            height = GetRealHeight(SlideShape)
            If height < Shortest Then Shortest = height
        Next
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            SetRealHeight SlideShape, Shortest
        Next SlideShape
        
    End If
    
End Sub

Sub ObjectsSizeToWidest(Optional ByVal control As IRibbonControl)
    Set MyDocument = Application.ActiveWindow
    Dim width       As Single
    Dim Widest      As Single
    Dim SlideShape  As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Widest = GetRealWidth(MyDocument.Selection.ChildShapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            width = GetRealWidth(SlideShape)
            If width > Widest Then Widest = width
        Next SlideShape
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            SetRealWidth SlideShape, Widest
        Next SlideShape
        
    Else
        
        Widest = GetRealWidth(MyDocument.Selection.ShapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            width = GetRealWidth(SlideShape)
            If width > Widest Then Widest = width
        Next SlideShape
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            SetRealWidth SlideShape, Widest
        Next SlideShape
        
    End If
    
End Sub

Sub ObjectsSizeToNarrowest(Optional ByVal control As IRibbonControl)
    Set MyDocument = Application.ActiveWindow
    Dim width       As Single
    Dim Narrowest   As Single
    Dim SlideShape  As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Narrowest = GetRealWidth(MyDocument.Selection.ChildShapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            width = GetRealWidth(SlideShape)
            If width < Narrowest Then Narrowest = width
        Next SlideShape
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            SetRealWidth SlideShape, Narrowest
        Next SlideShape
        
    Else
        
        Narrowest = GetRealWidth(MyDocument.Selection.ShapeRange(1))
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            width = GetRealWidth(SlideShape)
            If width < Narrowest Then Narrowest = width
        Next SlideShape
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            SetRealWidth SlideShape, Narrowest
        Next SlideShape
        
    End If
    
End Sub

Sub ObjectsSameHeight(Optional ByVal control As IRibbonControl)
    Set MyDocument = Application.ActiveWindow
    Dim height      As Single
    Dim SlideShape  As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        If MyDocument.Selection.HasChildShapeRange Then
            height = GetRealHeight(MyDocument.Selection.ChildShapeRange(1))
            
            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealHeight SlideShape, height
            Next SlideShape
        Else
            height = GetRealHeight(MyDocument.Selection.ShapeRange(1))
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                SetRealHeight SlideShape, height
            Next SlideShape
        End If
    Else
        If MyDocument.Selection.HasChildShapeRange Then
            height = GetRealHeight(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.count))
            
            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealHeight SlideShape, height
            Next SlideShape
        Else
            height = GetRealHeight(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.count))
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                SetRealHeight SlideShape, height
            Next SlideShape
        End If
    End If
End Sub

Sub ObjectsSameWidth(Optional ByVal control As IRibbonControl)
    Set MyDocument = Application.ActiveWindow
    Dim width       As Single
    Dim SlideShape  As shape
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        If MyDocument.Selection.HasChildShapeRange Then
            width = GetRealWidth(MyDocument.Selection.ChildShapeRange(1))
            
            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealWidth SlideShape, width
            Next SlideShape
        Else
            width = GetRealWidth(MyDocument.Selection.ShapeRange(1))
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                SetRealWidth SlideShape, width
            Next SlideShape
        End If
    Else
        If MyDocument.Selection.HasChildShapeRange Then
            width = GetRealWidth(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.count))
            
            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                SetRealWidth SlideShape, width
            Next SlideShape
        Else
            width = GetRealWidth(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.count))
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                SetRealWidth SlideShape, width
            Next SlideShape
        End If
    End If
End Sub

Sub ObjectsSameHeightAndWidth(Optional ByVal control As IRibbonControl)
    Set MyDocument = Application.ActiveWindow
    Dim height      As Single
    Dim width       As Single
    Dim SlideShape  As shape
    Dim sizeFailed  As Boolean
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    sizeFailed = False
    
    If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        If MyDocument.Selection.HasChildShapeRange Then
            height = GetRealHeight(MyDocument.Selection.ChildShapeRange(1))
            width = GetRealWidth(MyDocument.Selection.ChildShapeRange(1))
            
            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                If Not SetRealSize(SlideShape, width, height) Then sizeFailed = True
            Next SlideShape
        Else
            height = GetRealHeight(MyDocument.Selection.ShapeRange(1))
            width = GetRealWidth(MyDocument.Selection.ShapeRange(1))
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                If Not SetRealSize(SlideShape, width, height) Then sizeFailed = True
            Next SlideShape
        End If
    Else
        If MyDocument.Selection.HasChildShapeRange Then
            height = GetRealHeight(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.count))
            width = GetRealWidth(MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.count))
            
            For Each SlideShape In MyDocument.Selection.ChildShapeRange
                If Not SetRealSize(SlideShape, width, height) Then sizeFailed = True
            Next SlideShape
        Else
            height = GetRealHeight(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.count))
            width = GetRealWidth(MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.count))
            
            For Each SlideShape In MyDocument.Selection.ShapeRange
                If Not SetRealSize(SlideShape, width, height) Then sizeFailed = True
            Next SlideShape
        End If
    End If

    If sizeFailed Then
        MsgBox "Same Size could not match both visual dimensions for one or more shapes at their current rotation.", vbExclamation, "PPTK"
    End If
End Sub

' Swaps the stacking positions of two shapes and leaves every other shape where it was.
' The upper shape goes down d steps (to the lower slot, pushing the lower shape up one),
' then the lower shape goes up d-1 steps (to the upper slot).
Private Sub PPTK_SwapZOrder(ByVal first As Shape, ByVal second As Shape)
    Dim upper As Shape
    Dim lower As Shape
    Dim steps As Long
    Dim i As Long
    If first.ZOrderPosition > second.ZOrderPosition Then
        Set upper = first
        Set lower = second
    Else
        Set upper = second
        Set lower = first
    End If
    steps = upper.ZOrderPosition - lower.ZOrderPosition
    For i = 1 To steps
        upper.ZOrder msoSendBackward
    Next i
    For i = 1 To steps - 1
        lower.ZOrder msoBringForward
    Next i
End Sub

Sub ObjectsSwapPosition(Optional ByVal control As IRibbonControl)
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim Left1       As Single
    Dim Left2       As Single
    Dim Top1        As Single
    Dim Top2        As Single
    
    If ActiveWindow.Selection.ShapeRange.count = 2 Then
        
        Left1 = GetRealLeft(ActiveWindow.Selection.ShapeRange(1))
        Left2 = GetRealLeft(ActiveWindow.Selection.ShapeRange(2))
        Top1 = GetRealTop(ActiveWindow.Selection.ShapeRange(1))
        Top2 = GetRealTop(ActiveWindow.Selection.ShapeRange(2))
        
        SetRealLeft ActiveWindow.Selection.ShapeRange(1), Left2
        SetRealLeft ActiveWindow.Selection.ShapeRange(2), Left1
        SetRealTop ActiveWindow.Selection.ShapeRange(1), Top2
        SetRealTop ActiveWindow.Selection.ShapeRange(2), Top1
        
        
        PPTK_SwapZOrder ActiveWindow.Selection.ShapeRange(1), ActiveWindow.Selection.ShapeRange(2)
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.count = 2 Then
            
            Left1 = GetRealLeft(MyDocument.Selection.ChildShapeRange(1))
            Left2 = GetRealLeft(MyDocument.Selection.ChildShapeRange(2))
            Top1 = GetRealTop(MyDocument.Selection.ChildShapeRange(1))
            Top2 = GetRealTop(MyDocument.Selection.ChildShapeRange(2))
            
            SetRealLeft ActiveWindow.Selection.ChildShapeRange(1), Left2
            SetRealLeft ActiveWindow.Selection.ChildShapeRange(2), Left1
            SetRealTop ActiveWindow.Selection.ChildShapeRange(1), Top2
            SetRealTop ActiveWindow.Selection.ChildShapeRange(2), Top1
            
            
            PPTK_SwapZOrder ActiveWindow.Selection.ChildShapeRange(1), ActiveWindow.Selection.ChildShapeRange(2)
            
        Else
            
            MsgBox "Select two shapes To swap positions."
            
        End If
        
    Else
        
        MsgBox "Select two shapes To swap positions."
        
    End If
    
End Sub

Sub ObjectsSwapPositionCentered(Optional ByVal control As IRibbonControl)
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim Left1, Left2, Top1, Top2, Width1, Width2, Height1, Height2 As Single
    
    If ActiveWindow.Selection.ShapeRange.count = 2 Then
        
        Left1 = GetRealLeft(ActiveWindow.Selection.ShapeRange(1))
        Left2 = GetRealLeft(ActiveWindow.Selection.ShapeRange(2))
        Top1 = GetRealTop(ActiveWindow.Selection.ShapeRange(1))
        Top2 = GetRealTop(ActiveWindow.Selection.ShapeRange(2))
        Width1 = GetRealWidth(ActiveWindow.Selection.ShapeRange(1))
        Width2 = GetRealWidth(ActiveWindow.Selection.ShapeRange(2))
        Height1 = GetRealHeight(ActiveWindow.Selection.ShapeRange(1))
        Height2 = GetRealHeight(ActiveWindow.Selection.ShapeRange(2))
        
        SetRealLeft ActiveWindow.Selection.ShapeRange(1), (Left2 + (Width2 - Width1) / 2)
        SetRealLeft ActiveWindow.Selection.ShapeRange(2), (Left1 + (Width1 - Width2) / 2)
        SetRealTop ActiveWindow.Selection.ShapeRange(1), (Top2 + (Height2 - Height1) / 2)
        SetRealTop ActiveWindow.Selection.ShapeRange(2), (Top1 + (Height1 - Height2) / 2)

        
        PPTK_SwapZOrder ActiveWindow.Selection.ShapeRange(1), ActiveWindow.Selection.ShapeRange(2)
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.count = 2 Then
            
        Left1 = GetRealLeft(ActiveWindow.Selection.ChildShapeRange(1))
        Left2 = GetRealLeft(ActiveWindow.Selection.ChildShapeRange(2))
        Top1 = GetRealTop(ActiveWindow.Selection.ChildShapeRange(1))
        Top2 = GetRealTop(ActiveWindow.Selection.ChildShapeRange(2))
        Width1 = GetRealWidth(ActiveWindow.Selection.ChildShapeRange(1))
        Width2 = GetRealWidth(ActiveWindow.Selection.ChildShapeRange(2))
        Height1 = GetRealHeight(ActiveWindow.Selection.ChildShapeRange(1))
        Height2 = GetRealHeight(ActiveWindow.Selection.ChildShapeRange(2))
        
        SetRealLeft ActiveWindow.Selection.ChildShapeRange(1), (Left2 + (Width2 - Width1) / 2)
        SetRealLeft ActiveWindow.Selection.ChildShapeRange(2), (Left1 + (Width1 - Width2) / 2)
        SetRealTop ActiveWindow.Selection.ChildShapeRange(1), (Top2 + (Height2 - Height1) / 2)
        SetRealTop ActiveWindow.Selection.ChildShapeRange(2), (Top1 + (Height1 - Height2) / 2)

        
        PPTK_SwapZOrder ActiveWindow.Selection.ChildShapeRange(1), ActiveWindow.Selection.ChildShapeRange(2)
            
        Else
            
            MsgBox "Select two shapes To swap positions."
            
        End If
        
    Else
        
        MsgBox "Select two shapes To swap positions."
        
    End If
    
End Sub
