Attribute VB_Name = "ChangeFontSize"
Sub ChangeFontSize()
    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
    If shp.HasTextFrame = True Then
        If shp.TextFrame.HasText = True Then
            If shp.TextFrame.TextRange.Font.Size = 14 Then
                shp.TextFrame.TextRange.Font.Size = 11
            End If
            If shp.TextFrame.TextRange.Font.Size = 21 Then
                shp.TextFrame.TextRange.Font.Size = 16.5
            End If
            If shp.TextFrame.TextRange.Font.Size = 28 Then
                shp.TextFrame.TextRange.Font.Size = 22
            End If
        End If
    End If
    Next
    Next
End Sub
