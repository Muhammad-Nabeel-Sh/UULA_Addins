Attribute VB_Name = "Single_space"
Sub set_gap()
Dim sngGap As Single
Dim rayShapes() As Shape
Dim L As Long
On Error Resume Next
If ActiveWindow.Selection.ShapeRange.Count < 2 Then
MsgBox "Select at least 2 shapes"
Exit Sub
End If
On Error GoTo 0
ReDim rayShapes(1 To ActiveWindow.Selection.ShapeRange.Count)
sngGap = in2Points(0.093)
For L = 1 To ActiveWindow.Selection.ShapeRange.Count
Set rayShapes(L) = ActiveWindow.Selection.ShapeRange(L)
Next L
' make sure selected shapes are sorted by Top value
Call SortByLeft(rayShapes)
' set the gap
For L = 2 To UBound(rayShapes)
Debug.Print rayShapes(L).Name
rayShapes(L).Top = rayShapes(L - 1).Top + rayShapes(L - 1).Height + sngGap
Next L
End Sub

Sub SortByLeft(Arrayin As Variant)
' sort the shapes based on their Top value
Dim b_Cont As Boolean
Dim lngCount As Long
Dim vSwap As Shape
Do
    b_Cont = False
    For lngCount = LBound(Arrayin) To UBound(Arrayin) - 1
        Debug.Print Arrayin(lngCount).Name
        If Arrayin(lngCount).Top > Arrayin(lngCount + 1).Top Then
            Set vSwap = Arrayin(lngCount)
            Set Arrayin(lngCount) = Arrayin(lngCount + 1)
            Set Arrayin(lngCount + 1) = vSwap
            b_Cont = True
        End If
    Next lngCount
Loop Until Not b_Cont
'release objects
Set vSwap = Nothing
End Sub

Function in2Points(inVal As Single) As Single
'convert inches to points
in2Points = inVal * 72
End Function
