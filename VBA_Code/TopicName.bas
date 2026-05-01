Attribute VB_Name = "TopicName"
Option Explicit
Sub TopicName()
Dim shp As Shape
For Each shp In ActiveWindow.Selection.ShapeRange
            On Error Resume Next
With ActiveWindow.Selection
    .TextRange.Font.Color = RGB(0, 0, 0)
    .TextRange.Font.Name = "UULA Sans"
    .TextRange.Font.NameComplexScript = "UULA Sans"
    .TextRange.Font.Size = 16.5
    .TextRange.Font.Bold = msoTrue
    
    .TextRange2.ParagraphFormat.LeftIndent = 0
    .TextRange2.ParagraphFormat.FirstLineIndent = 0
    .TextRange2.ParagraphFormat.SpaceAfter = 0
    .TextRange2.ParagraphFormat.SpaceBefore = 0
    '.TextRange2.ParagraphFormat.LineRuleWithin = msoTrue
    .TextRange2.ParagraphFormat.SpaceWithin = 1

    .TextRange2.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
    .TextRange2.ParagraphFormat.Alignment = msoAlignRight
    
    .TextRange.ParagraphFormat.Bullet.Type = ppBulletNone
    
    .ShapeRange.TextFrame2.MarginBottom = 0
    .ShapeRange.TextFrame2.MarginLeft = 0
    .ShapeRange.TextFrame2.MarginRight = 0
    .ShapeRange.TextFrame2.MarginTop = 0
    .ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorNone
    .ShapeRange.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    .ShapeRange.Width = (72 * 6.5)
    .ShapeRange.Align msoAlignCenters, msoCTrue
End With
Next shp
End Sub


