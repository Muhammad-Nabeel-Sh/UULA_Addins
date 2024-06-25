Attribute VB_Name = "CustomNumberSlides"
Sub CustomNumberSlides()
    Dim oSl As Slide
    Dim oSh As Shape
    Dim x As Long
    Dim oNumber As Shape
    Dim lStartNumberingOn As Long
    Dim lStartingNumber As Long
    ' Dim sFormat As String

    ' What slide should the numbering start ON
    lStartNumberingOn = InputBox("Enter the first actual slide: ")
    ' What should the first slide number BE
    lStartingNumber = 1
    ' How should the number be displayed
    'sFormat = "00"   ' three digits, add leading zeros

    ' First, delete any previous numbers
'    For Each oSl In ActivePresentation.Slides
'        For x = oSl.Shapes.Count To 1 Step -1
'        ' Wild card for searching for slide number Textboxes
'            If oSl.Shapes(x).Name Like "Slide Number Placeholder *" Then
'               oSl.Shapes(x).Delete
'            End If
'            If oSl.Shapes(x).Tags("MyNumber") = "Y" Then
'               oSl.Shapes(x).Delete
'            End If
'        Next
'    Next

    ' Now add new slide numbers:
    For x = lStartNumberingOn To ActivePresentation.Slides.Count - 1
        Set oSl = ActivePresentation.Slides(x)
        With oSl.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 350, 90)

            ' Add a tag so we can find and delete the number later
            .Tags.Add "MyNumber", "Y"
            ' Set the text of the number
            .TextFrame.TextRange.Text = Format(lStartingNumber)

            ' Edit these to change the position of the number
            .Left = 248
            .Height = 20
            .Top = 745
            .Width = 50

            ' Edit these to change the formatting of the number
            With .TextFrame.TextRange
                .ParagraphFormat.Alignment = ppAlignCenter
                .Font.Name = "UULA Sans"
                .Font.Size = 12
                .Font.Color.RGB = RGB(166, 166, 166)
            End With
        End With

        lStartingNumber = lStartingNumber + 1

    Next

End Sub
