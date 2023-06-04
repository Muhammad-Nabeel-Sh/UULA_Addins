Attribute VB_Name = "FontChange"
Option Explicit

Sub FontChange()
    
    Dim sld As Slide
    Dim shp As Shape
    Dim subshp As Shape
    Dim i As Long
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then        ' Not all shapes do
            If shp.TextFrame.HasText Then        ' the shape may contain no text
            shp.TextFrame.TextRange.Font.Name = "UULA Sans"
            shp.TextFrame.TextRange.Font.NameComplexScript = "UULA Sans"
        End If
    End If
    
    If shp.Type = msoGroup Then
    For Each subshp In shp.GroupItems
        If subshp.HasTextFrame Then        ' Not all shapes do
                If subshp.TextFrame.HasText Then          ' the shape may contain no text
                   subshp.TextFrame.TextRange.Font.Name = "UULA Sans"
                   subshp.TextFrame.TextRange.Font.NameComplexScript = "UULA Sans"
                End If
        End If
    Next subshp
    End If

        If shp.Type = msoPlaceholder Then
        If shp.HasTextFrame Then        ' Not all shapes do
                   shp.TextFrame.TextRange.Font.Name = "UULA Sans"
                   shp.TextFrame.TextRange.Font.NameComplexScript = "UULA Sans"
        End If
    End If
    
    If shp.Type = msoTextBox Then
        If shp.HasTextFrame Then        ' Not all shapes do
                   shp.TextFrame.TextRange.Font.Name = "UULA Sans"
                   shp.TextFrame.TextRange.Font.NameComplexScript = "UULA Sans"
        End If
    End If
        
    Next shp
    Next sld
End Sub