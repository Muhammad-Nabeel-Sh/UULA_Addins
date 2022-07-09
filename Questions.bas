Attribute VB_Name = "Questions"
Option Explicit
Sub Questions()
    Dim shp As Shape
    For Each shp In ActiveWindow.Selection.ShapeRange
                 On Error Resume Next
        With ActiveWindow.Selection
            .TextRange.Font.Color = RGB(0, 0, 0)
            .TextRange.Font.Name = "Avenir Next Arabic"
            .TextRange.Font.NameComplexScript = "Avenir Next Arabic"
            .TextRange.Font.Size = 11
            .TextRange.Font.Bold = msoFalse
            
            .TextRange.ParagraphFormat.Bullet.Font.Name = "Avenir Next Arabic Black"
            .TextRange.ParagraphFormat.Bullet.Character = 81
            .TextRange.ParagraphFormat.Bullet.RelativeSize = 1
            .TextRange.ParagraphFormat.Bullet.Font.Color = RGB(255, 0, 0)
            
            .TextRange2.ParagraphFormat.FirstLineIndent = -(72 * 0.2)
            .TextRange2.ParagraphFormat.LeftIndent = 72 * 0.2
            '.TextRange2.ParagraphFormat.LineRuleAfter = False
            '.TextRange2.ParagraphFormat.LineRuleBefore = False
            .TextRange2.ParagraphFormat.SpaceAfter = 0
            .TextRange2.ParagraphFormat.SpaceBefore = 0
            .TextRange2.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
            .TextRange2.ParagraphFormat.Alignment = msoAlignRight
            '.TextRange.ParagraphFormat.LineRuleWithin = True
            .TextRange.ParagraphFormat.SpaceWithin = 1
            
            .ShapeRange.TextFrame2.MarginBottom = 0
            .ShapeRange.TextFrame2.MarginLeft = 0
            .ShapeRange.TextFrame2.MarginRight = 0
            .ShapeRange.TextFrame2.MarginTop = 0
            .ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorNone
            .ShapeRange.TextFrame2.VerticalAnchor = msoAnchorTop
            .ShapeRange.TextFrame.AutoSize = ppAutoSizeShapeToFitText
            .ShapeRange.Width = (72 * 6.5)
            .ShapeRange.Align msoAlignCenters, msoCTrue
        End With
    Next shp
End Sub
