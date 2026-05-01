Attribute VB_Name = "RemoveNewNum"
Sub RemoveNewNum()
    Dim oSl As Slide
    Dim oSh As Shape
    Dim x As Long
    
    For Each oSl In ActivePresentation.Slides
        For x = oSl.Shapes.Count To 1 Step -1
            If oSl.Shapes(x).Tags("MyNumber") = "Y" Then
               oSl.Shapes(x).Delete
            End If
        Next
    Next
    
End Sub

