VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Seguimiento 
   Caption         =   "UserForm1"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11805
   OleObjectBlob   =   "Seguimiento.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Seguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Actualizar el registro
Private Sub CommandButton1_Click()

For i = 1 To 11
    Sheets("Hoja2").Activate
    ActiveSheet.Cells(1, i).Select
    ActiveCell.Offset(0, i - 1).Value = Me.Controls("TextBox" & i).Value
Next i
Unload Me
End Sub
'
'Cerrar formulario
Private Sub CommandButton2_Click()
Unload Me
End Sub
'
'Llenar los cuadro de texto con los datos del registro elegido
Private Sub UserForm_Initialize()
For i = 1 To 11
    Sheets("Hoja2").Activate
    ActiveSheet.Cells(1, i).Select
    Me.Controls("TextBox" & i).Value = ActiveCell.Offset(0, i - 1).Value
Next i
End Sub


