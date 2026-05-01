Option Explicit
Sub NormalText()
Dim shp As Shape
Dim i As Integer
For Each shp In ActiveWindow.Selection.ShapeRange
            On Error Resume Next
If shp.HasTextFrame Then
If shp.TextFrame.HasText Then

'shp.TextFrame.TextRange.Select
'Application.CommandBars.ExecuteMso ("EquationInsertNew")
'shp.TextFrame.TextRange.InsertAfter (" ")
'Application.CommandBars.ExecuteMso ("AlignLeft")

For i = 1 To shp.TextFrame.TextRange.Paragraphs.Count
    shp.TextFrame.TextRange.Paragraphs(i).Select
    Application.CommandBars.ExecuteMso ("EquationInsertNew")
    shp.TextFrame.TextRange.InsertAfter (" ")
    Application.CommandBars.ExecuteMso ("AlignLeft")
 '  shp.TextFrame.TextRange.Lines
 '  shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
Next i

End If
End If

Next shp
End Sub