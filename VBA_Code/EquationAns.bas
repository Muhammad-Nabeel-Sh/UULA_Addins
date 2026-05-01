Attribute VB_Name = "EquationAns"
Option Explicit
Sub EquationAns()
Dim shp As Shape
For Each shp In ActiveWindow.Selection.ShapeRange
            On Error Resume Next
With ActiveWindow.Selection
    .TextRange.Font.Color = RGB(255, 0, 255)
    .TextRange.Font.Name = "UULA Sans"
    .TextRange.Font.NameComplexScript = "UULA Sans"
    .TextRange.Font.Size = 11
    .TextRange.Font.Bold = msoTrue
    .TextRange.ParagraphFormat.Bullet.Type = ppBulletNone
End With
Next shp
End Sub


