Sub SetTableFont()
    Dim oShp As Shape
    Dim oTbl As Table
    Dim oSld As Slide
    Dim I As Long
    Dim J As Long
    
    For Each oSld In ActivePresentation.Slides
        For Each oShp In oSld.Shapes
            If oShp.HasTable Then
                Set oTbl = oShp.Table
                For I = 1 To oTbl.Columns.Count
                    For J = 1 To oTbl.Rows.Count
                        With oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font
                            .Size = 11
                            .Name = "Avenir Next Arabic"
                            .NameComplexScript = "Avenir Next Arabic"
                        End With
                        If oTbl.FirstCol = TRUE Then
                            oTbl.Cell(J, 1).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        End If
                        If oTbl.FirstRow = TRUE Then
                            oTbl.Cell(1, I).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        End If
                    Next J
                Next I
                oShp.Width = (72 * 6.5)
                oShp.Left = 72 * 0.5
            End If
        Next oShp
    Next oSld
End Sub