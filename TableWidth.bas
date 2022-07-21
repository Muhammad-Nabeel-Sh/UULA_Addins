Attribute VB_Name = "TableWidth"
Sub TableWidth()
    Dim oShp As Shape
    Dim oTbl As Table
    Dim oSld As Slide
    Dim I As Long
    Dim J As Long
    
    For Each oSld In ActivePresentation.Slides
        For Each oShp In oSld.Shapes
            If oShp.HasTable Then
                Set oTbl = oShp.Table
                oShp.Width = (72 * 6.5)
                oShp.Left = 72 * 0.5
            End If
        Next oShp
    Next oSld
End Sub