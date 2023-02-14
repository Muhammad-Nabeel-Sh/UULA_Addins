Attribute VB_Name = "QuestionsEN"
Option Explicit
Sub QuestionsEN()
    Dim shp As Shape
    For Each shp In ActiveWindow.Selection.ShapeRange
                     On Error Resume Next
        With ActiveWindow.Selection
            .TextRange.Font.Color = RGB(0, 0, 0)
            .TextRange.Font.Name = "UULA Sans"
            .TextRange.Font.NameComplexScript = "UULA Sans"
            .TextRange.Font.Size = 11
            .TextRange.Font.Bold = msoFalse
            
            .TextRange.ParagraphFormat.Bullet.Font.Name = "UULA Sans Black"
            .TextRange.ParagraphFormat.Bullet.Character = 81
            .TextRange.ParagraphFormat.Bullet.RelativeSize = 1
            .TextRange.ParagraphFormat.Bullet.Font.Color = RGB(255, 0, 0)
            
            .TextRange2.ParagraphFormat.FirstLineIndent = -(72 * 0.2)
            .TextRange2.ParagraphFormat.LeftIndent = 72 * 0.2
            .TextRange2.ParagraphFormat.SpaceAfter = 0
            .TextRange2.ParagraphFormat.SpaceBefore = 0
            .TextRange2.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight
            .TextRange2.ParagraphFormat.Alignment = msoAlignLeft
            '.TextRange2.ParagraphFormat.LineRuleWithin = msoTrue
            .TextRange2.ParagraphFormat.SpaceWithin = 1
            
            .ShapeRange.TextFrame2.MarginBottom = 0
            .ShapeRange.TextFrame2.MarginLeft = 0
            .ShapeRange.TextFrame2.MarginRight = 0
            .ShapeRange.TextFrame2.MarginTop = 0
            .ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorNone
            .ShapeRange.TextFrame2.VerticalAnchor = msoAnchorTop
            '.ShapeRange.TextFrame.AutoSize = ppAutoSizeShapeToFitText
            .ShapeRange.Width = (72 * 6.5)
            .ShapeRange.Align msoAlignCenters, msoCTrue
        End With
    Next shp
End Sub
