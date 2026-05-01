VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TableStyle 
   Caption         =   "Table Styling Options"
   ClientHeight    =   3445
   ClientLeft      =   104
   ClientTop       =   429
   ClientWidth     =   7280
   OleObjectBlob   =   "TableStyle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TableStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox8_Click()

End Sub

Private Sub CommandButton1_Click()
    ' Hide the dialog box while it works
    Me.Hide
    
    ' Call the main macro and pass the True/False values from your checkboxes
    Call StyleSelectedTableX( _
        CheckBox1.Value, _
        CheckBox8.Value, _
        CheckBox9.Value, _
        CheckBox12.Value, _
        CheckBox5.Value, _
        CheckBox10.Value, _
        CheckBox11.Value)
        
    ' Unload the form from memory
    Unload Me
End Sub

