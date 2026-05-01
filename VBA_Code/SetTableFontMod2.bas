Attribute VB_Name = "SetTableFontMod2"
Sub SetTableFontMod2()
    Dim oShp As Shape
    Dim oTbl As Table
    Dim oSld As Slide
    Dim I As Long
    Dim J As Long
    Dim X As Integer
    
    For Each oSld In ActivePresentation.Slides
        For Each oShp In oSld.Shapes
            If oShp.HasTable Then
                Set oTbl = oShp.Table
                For I = 1 To oTbl.Columns.Count
                    For J = 1 To oTbl.Rows.Count
                        With oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font
                            .Size = 11
                            .Name = "UULA Sans"
                            .NameComplexScript = "UULA Sans"
                        End With
                        ' If oTbl.FirstCol = TRUE Then
                        '     oTbl.Cell(J, 1).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        ' End If
                        ' If oTbl.FirstRow = TRUE Then
                        '     oTbl.Cell(1, I).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        ' End If
                        'If (oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color <> RGB(31, 113, 222)) And (oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color <> RGB(255, 0, 0)) And (oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color <> RGB(31, 113, 221)) And (oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color <> 0) And (oTbl.Cell(J, I).Shape.TextFrame2.TextRange.Font.Fill.ForeColor <> RGB(31, 113, 222)) And (oTbl.Cell(J, I).Shape.TextFrame2.TextRange.Font.Fill.ForeColor <> RGB(255, 0, 0)) And (oTbl.Cell(J, I).Shape.TextFrame2.TextRange.Font.Fill.ForeColor <> RGB(31, 113, 221)) And (oTbl.Cell(J, I).Shape.TextFrame2.TextRange.Font.Fill.ForeColor <> 0) Then
                        For X = 0 To 255
                        If oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(X, X, X) Then
                            oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        End If
                        Next X
                    Next J
                Next I
                ' oShp.Width = (72 * 6.5)
                ' oShp.Left = 72 * 0.5
            End If
        Next oShp
    Next oSld
End Sub
