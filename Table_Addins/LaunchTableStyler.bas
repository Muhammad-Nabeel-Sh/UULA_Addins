Attribute VB_Name = "LaunchTableStyler"
' Run THIS macro to open your dialog box
Sub LaunchTableStyler()
    ' 1. Check if a shape is selected BEFORE opening the dialog box
    If ActiveWindow.Selection.Type = ppSelectionNone Then
        MsgBox "Please select a table first before opening the styler.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' 2. Check if the selected shape is actually a table
    If ActiveWindow.Selection.ShapeRange(1).HasTable = msoFalse Then
        MsgBox "The selected object is not a table. Please select a table.", vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    ' Open the UserForm
    TableStyle.Show
End Sub

' This macro is now triggered by your UserForm button
Sub StyleSelectedTableX(pSuperHeaderRow As Boolean, pHeaderRow As Boolean, _
                       pTotalRow As Boolean, pBandedRows As Boolean, _
                       pFirstColumn As Boolean, pLastColumn As Boolean, _
                       pBandedColumns As Boolean)

    ' =====================================================================
    ' PREDEFINED PARAMETERS (Only Colors, Sizes, and Fonts remain here)
    ' =====================================================================
    
    Const pApplyTableWidth As Boolean = True
    Const pTableWidth As Single = 468
    Const pCellHeight As Single = 20.16
    Const pApplyCellWidth As Boolean = False
    Const pCellWidth As Single = 100
    
    Const pApplyFont As Boolean = True
    Const pFontName As String = "UULA Sans"
    Const pFontNameComplexScript = "UULA Sans"
    Const pFontSize As Single = 11
    
    Const pApplyCustomShading As Boolean = True
    Const pApplyCustomBorders As Boolean = True
    Const pBorderWeight As Single = 0.5
    
    ' =====================================================================
    
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Integer, c As Integer
    Dim mainHeaderRowIndex As Integer
    
    Dim myRGBSuperHeader As Long
    Dim myRGBHeader As Long
    Dim myRGBShading1 As Long
    Dim myRGBShading2 As Long
    Dim myRGBFirstColumn As Long
    Dim myRGBBorder As Long
    
    myRGBSuperHeader = RGB(163, 176, 193)
    myRGBHeader = RGB(202, 208, 216)
    myRGBFirstColumn = RGB(218, 224, 233)
    myRGBShading1 = RGB(255, 255, 255)
    myRGBShading2 = RGB(241, 241, 241)
    myRGBBorder = RGB(217, 217, 217)

    ' Set references
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    Set tbl = shp.Table
    
    ' Apply Table Structure Toggles (From UserForm)
    tbl.FirstRow = pHeaderRow
    tbl.LastRow = pTotalRow
    tbl.HorizBanding = pBandedRows
    tbl.FirstCol = pFirstColumn
    tbl.LastCol = pLastColumn
    tbl.VertBanding = pBandedColumns
    
    ' Determine Main Header position
    If pSuperHeaderRow Then
        mainHeaderRowIndex = 2
    Else
        mainHeaderRowIndex = 1
    End If
    
    ' Apply Table Width
    If pApplyTableWidth Then shp.Width = pTableWidth
    
    ' Apply Cell Sizing
    For r = 1 To tbl.Rows.Count
        tbl.Rows(r).Height = pCellHeight
    Next r
    
    If pApplyCellWidth Then
        For c = 1 To tbl.Columns.Count
            tbl.Columns(c).Width = pCellWidth
        Next c
    End If
    
    ' Loop through cells
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            
            ' Apply Custom Shading
            If pApplyCustomShading Then
                With tbl.Cell(r, c).Shape.Fill
                    .Visible = msoTrue
                    .Solid
                    
                    If pSuperHeaderRow And r = 1 Then
                        .ForeColor.RGB = myRGBSuperHeader
                    ElseIf pHeaderRow And r = mainHeaderRowIndex Then
                        .ForeColor.RGB = myRGBHeader
                    ElseIf pFirstColumn And c = 1 Then
                        .ForeColor.RGB = myRGBFirstColumn
                    ElseIf pBandedRows Then
                        If r Mod 2 = 0 Then
                            .ForeColor.RGB = myRGBShading1
                        Else
                            .ForeColor.RGB = myRGBShading2
                        End If
                    Else
                        .ForeColor.RGB = myRGBShading1
                    End If
                End With
            End If
            
            ' Apply Custom Borders
            If pApplyCustomBorders Then
                With tbl.Cell(r, c)
                    .Borders(ppBorderTop).Visible = msoTrue
                    .Borders(ppBorderTop).ForeColor.RGB = myRGBBorder
                    .Borders(ppBorderTop).Weight = pBorderWeight
                    
                    .Borders(ppBorderBottom).Visible = msoTrue
                    .Borders(ppBorderBottom).ForeColor.RGB = myRGBBorder
                    .Borders(ppBorderBottom).Weight = pBorderWeight
                    
                    .Borders(ppBorderLeft).Visible = msoTrue
                    .Borders(ppBorderLeft).ForeColor.RGB = myRGBBorder
                    .Borders(ppBorderLeft).Weight = pBorderWeight
                    
                    .Borders(ppBorderRight).Visible = msoTrue
                    .Borders(ppBorderRight).ForeColor.RGB = myRGBBorder
                    .Borders(ppBorderRight).Weight = pBorderWeight
                End With
            End If
            
            ' Apply Font Styles
            If pApplyFont Then
                If tbl.Cell(r, c).Shape.HasTextFrame Then
                    With tbl.Cell(r, c).Shape.TextFrame.TextRange.Font
                        .Name = pFontName
                        .NameComplexScript = pFontNameComplexScript
                        .Size = pFontSize
                        .Color = RGB(0, 0, 0)
                        
                        If (pSuperHeaderRow And r = 1) Or (pHeaderRow And r = mainHeaderRowIndex) Or (pFirstColumn And c = 1) Then
                            .Bold = msoTrue
                        Else
                            .Bold = msoFalse
                        End If
                    End With
                    
                    With tbl.Cell(r, c).Shape.TextFrame
                        .VerticalAnchor = msoAnchorMiddle ' Center Vertically
                        .TextRange.ParagraphFormat.Alignment = ppAlignCenter ' Center Horizontally
                    End With
                    
                End If
            End If
            
        Next c
    Next r

    MsgBox "Table styling applied successfully!", vbInformation, "Done"

End Sub

