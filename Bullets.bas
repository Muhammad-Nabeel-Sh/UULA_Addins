Attribute VB_Name = "Bullets"
Option Explicit
Sub Bullets()
    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
    If shp.HasTextFrame = True Then
        If shp.TextFrame.HasText = True Then
            If shp.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = True Then
                shp.TextFrame2.TextRange.ParagraphFormat.LeftIndent = 72 * 0.2
                shp.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = -(72 * 0.2)
            End If
            If shp.TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Cairo Black" Then
                With shp.TextFrame.TextRange
                .ParagraphFormat.Bullet.Font.Name = "UULA Sans Black"
                .ParagraphFormat.Bullet.Character = 81
                End With
            End If
            If shp.TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Avenir Next Arabic Black" Then
                With shp.TextFrame.TextRange
                .ParagraphFormat.Bullet.Font.Name = "UULA Sans Black"
                .ParagraphFormat.Bullet.Character = 81
            End With
            End If
            If shp.TextFrame.TextRange.ParagraphFormat.Bullet.Type = ppBulletPicture Then
                With shp.TextFrame.TextRange
                .ParagraphFormat.Bullet.Font.Name = "UULA Sans Black"
                .ParagraphFormat.Bullet.Character = 79
                End With
            End If
            If shp.TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Wingdings 2" And shp.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 153 Then
                With shp.TextFrame.TextRange
                .ParagraphFormat.Bullet.Font.Name = "UULA Sans Black"
                .ParagraphFormat.Bullet.Character = 79
                End With
            End If
            If shp.TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Wingdings 2" And shp.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 129 Then
                With shp.TextFrame.TextRange
                .ParagraphFormat.Bullet.Font.Name = "UULA Sans Black"
                .ParagraphFormat.Bullet.Character = 79
                End With
            End If
            If shp.TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Wingdings" And shp.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 161 Then
                With shp.TextFrame.TextRange
                .ParagraphFormat.Bullet.Font.Name = "UULA Sans Black"
                .ParagraphFormat.Bullet.Character = 79
                End With
            End If
        End If
    End If
    Next
    Next
End Sub


