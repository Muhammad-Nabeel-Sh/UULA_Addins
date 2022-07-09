Attribute VB_Name = "UnitColor"
Option Explicit

Sub UnitColor()
    
    Dim sld As Slide
    Dim shp As Shape
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then        ' Not all shapes do
            If shp.TextFrame.HasText Then        ' the shape may contain no text
            If shp.TextFrame.TextRange.Font.Color = RGB(42, 201, 222) Then
                shp.TextFrame.TextRange.Font.Color = RGB(166, 166, 166)
            End If
        End If
    End If
Next shp
Next sld
End Sub
