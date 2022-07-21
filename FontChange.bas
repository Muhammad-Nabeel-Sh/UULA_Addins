Attribute VB_Name = "FontChange"
Option Explicit

Sub FontChange()
    
    Dim sld As Slide
    Dim shp As Shape
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then        ' Not all shapes do
            If shp.TextFrame.HasText Then        ' the shape may contain no text
            shp.TextFrame.TextRange.Font.Name = "Avenir Next Arabic"
            shp.TextFrame.TextRange.Font.NameComplexScript = "Avenir Next Arabic"
        End If
    End If
Next shp
Next sld
End Sub
