Attribute VB_Name = "MergeTextBoxes"
Sub MergeTextBoxes()
' This will merge the text from all selected text boxes into the first selected box then delete the other text boxes

    Dim oRng As ShapeRange
    Dim oFirstShape As Shape
    Dim oSh As Shape
    Dim x As Long

    Set oRng = ActiveWindow.Selection.ShapeRange
        Set oFirstShape = oRng(1)
        oFirstShape.TextFrame.TextRange.Text = _
            oFirstShape.TextFrame.TextRange.Text & vbCrLf

    For x = 2 To oRng.Count
        oFirstShape.TextFrame.TextRange.Text = _
            oFirstShape.TextFrame.TextRange.Text _
            & oRng(x).TextFrame.TextRange.Text
        If x < oRng.Count Then
            oFirstShape.TextFrame.TextRange.Text = _
                oFirstShape.TextFrame.TextRange.Text _
                & vbCrLf
        End If
    Next

    For x = oRng.Count To 2 Step -1
        oRng(x).Delete
    Next

    Set oRng = Nothing
    Set oFirstShape = Nothing
    Set oSh = Nothing

End Sub
