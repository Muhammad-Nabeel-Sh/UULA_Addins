Attribute VB_Name = "CompleteAns"
Option Explicit
Sub CompleteAns()
    Dim shp As Shape
    For Each shp In ActiveWindow.Selection.ShapeRange
                     On Error Resume Next
        With ActiveWindow.Selection
            .TextRange.Font.Color = RGB(31, 113, 222)
            .TextRange.Font.Name = "UULA Sans"
            .TextRange.Font.NameComplexScript = "UULA Sans"
            .TextRange.Font.Size = 11
            .TextRange.Font.Bold = msoTrue
            
            .TextRange2.ParagraphFormat.LeftIndent = 0
            .TextRange2.ParagraphFormat.FirstLineIndent = 0
            .TextRange2.ParagraphFormat.SpaceAfter = 0
            .TextRange2.ParagraphFormat.SpaceBefore = 0
            .TextRange2.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
            .TextRange2.ParagraphFormat.Alignment = msoAlignRight
            
            '.TextRange2.ParagraphFormat.LineRuleWithin = msoTrue
            .TextRange2.ParagraphFormat.SpaceWithin = 1
            
            .TextRange.ParagraphFormat.Bullet.Type = ppBulletNone
            
            .ShapeRange.TextFrame2.MarginBottom = 0
            .ShapeRange.TextFrame2.MarginLeft = 0
            .ShapeRange.TextFrame2.MarginRight = 0
            .ShapeRange.TextFrame2.MarginTop = 0
            .ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorNone
            .ShapeRange.TextFrame2.VerticalAnchor = msoAnchorTop
            
            .ShapeRange.TextFrame.AutoSize = ppAutoSizeShapeToFitText
            .ShapeRange.TextFrame.WordWrap = msoFalse
            
        End With
    Next shp
End Sub
