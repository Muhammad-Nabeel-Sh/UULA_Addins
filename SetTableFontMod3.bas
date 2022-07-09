Attribute VB_Name = "SetTableFontMod3"
Sub SetTableFontMod3()
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
                            .Name = "Avenir Next Arabic"
                            .NameComplexScript = "Avenir Next Arabic"
                        End With

                        If oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(24, 23, 23) Then
                            oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        End If
                        
                        If oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(59, 56, 56) Then
                            oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        End If
                        
                        If oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(118, 113, 113) Then
                            oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        End If
                        
                        If oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(175, 171, 171) Then
                            oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        End If
                        
                        If oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(208, 206, 206) Then
                            oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        End If
                        
                        If oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(231, 230, 230) Then
                            oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                        End If
                        
                        For X = 0 To 255
                        If (oTbl.Cell(J, I).Shape.TextFrame.TextRange.Font.Color = RGB(X, X, X)) Then
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
