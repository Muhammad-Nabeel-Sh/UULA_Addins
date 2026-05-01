Attribute VB_Name = "RemoveOldNum"
Sub RemoveOldNum()
    Dim oSl As Slide
    Dim oSh As Shape
    Dim x As Long
    
    For Each oSl In ActivePresentation.Slides
        For x = oSl.Shapes.Count To 1 Step -1
        ' Wild card for searching for slide number Textboxes
            If oSl.Shapes(x).Name Like "Slide Number Placeholder *" Then
               oSl.Shapes(x).Delete
            End If
        Next
    Next
    
End Sub
